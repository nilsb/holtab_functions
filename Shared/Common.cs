using Shared.Models;
using Microsoft.Extensions.Logging;
using RE = System.Text.RegularExpressions;
using Microsoft.Graph.Models;
using System.Threading.Channels;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace Shared
{
    public class Common
    {
        private readonly ILogger? log;
        private readonly Graph? msGraph;
        private readonly Settings? settings;
        private readonly string? CDNTeamID;
        private readonly string? cdnSiteId;
        private readonly string? SqlConnectionString;
        private readonly Services? services;
        public readonly string? CDNGroup;
        public readonly string quoteRegex = @"^([A-Z]?\d+)";
        public readonly string customerRegex = @"(\d+)(\s?|-?)";
        public readonly string orderRegex = @"B\d{6}|T\d{5}|A\d{5}|Z\d{5}|G\d{4}|R\d{2}|E\d{5,7}|F\d{5}|H\d{5,6}|K\d{5,6}|\d{5}-\d{2}|Q\d{5}-\d{2}";
        public List<char> illegalChars = new List<char>() { '~', '`', '!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '_', '+', '=', '{', '}', '|', '[', ']', '\\', ':', '\"', ';', '\'', '<', '>', ',', '.', '?', '/', 'å', 'ä', 'ö', 'Å', 'Ä', 'Ö', ' ', 'Ø', 'Æ', 'æ', 'ø', 'ü', 'Ü', 'µ', 'ẞ', 'ß' };

        public Common(Settings _settings, Graph _msGraph, bool debug)
        {
            log = _settings.log;
            msGraph = _msGraph;
            settings = _settings;

            if (settings != null)
            {
                SqlConnectionString = settings.SqlConnectionString;
                CDNTeamID = settings.CDNTeamID;
                cdnSiteId = settings.cdnSiteId;

                if (!string.IsNullOrEmpty(SqlConnectionString))
                {
                    services = new Services(SqlConnectionString, log);
                }

                if (!string.IsNullOrEmpty(CDNTeamID))
                {
                    FindGroupResult? findGroup = msGraph?.GetGroupById(CDNTeamID, debug).Result;

                    if (findGroup?.Success == true)
                    {
                        CDNGroup = findGroup.group;
                    }
                }
            }
        }

        public Order? GetOrderFromCDN(string orderNo, bool debug)
        {
            Order? returnValue = null;

            try
            {
                if(debug)
                    log?.LogInformation($"GetOrderFromCDN: Trying to find order with id {orderNo}.");

                returnValue = services?.GetOrderFromDB(orderNo, debug);

                if(returnValue != null)
                {
                    if(debug)
                        log?.LogInformation($"GetOrderFromCDN: Found order with id {orderNo}. Trying to fetch customer.");

                    FindCustomerResult customers = GetCustomer(returnValue, debug);

                    if (returnValue.Customer == null && customers.Success && customers.customer != null)
                    {
                        if(debug)
                            log?.LogInformation($"GetOrderFromCDN: Found customer {customers.customer.Name} for order with id {orderNo}.");

                        returnValue.CustomerName = customers.customer.Name;
                        returnValue.Customer = customers.customer;

                        if (customers.customer.ID != Guid.Empty)
                        {
                            returnValue.CustomerID = customers.customer.ID;
                        }
                    }

                    returnValue.ExternalId = orderNo;
                    UpdateOrder(returnValue, "customer info", debug);
                }
            }
            catch (Exception ex)
            {
                log?.LogError(ex.ToString());
                
                if(debug)
                    log?.LogInformation($"GetOrderFromCDN: Failed to get order {orderNo} from CDN with error: " + ex.ToString());
            }

            return returnValue;
        }

        public void UpdateOrder(Order order, string logmsg, bool debug)
        {
            if (order != null)
            {
                if(debug)
                    log?.LogTrace($"Updating {logmsg} for order {order.ExternalId} in database.");
                services?.UpdateOrderInDB(order);
            }
            else if(debug)
            {
                log?.LogTrace($"Failed to update {logmsg} for order in database: order object was null");
            }
        }

        public Order? UpdateOrCreateDbOrder(Order order, bool debug)
        {
            Order? returnValue = null;
            Order? DBOrder = null;

            if(order != null)
            {
                if(debug)
                    log?.LogInformation($"UpdateOrCreateDbOrder: Processing order {order.ExternalId} for CreateOrUpdateDB.");

                //try to find the customer in the order object sent as parameter (in case it changed or the order is new)
                FindCustomerResult DBCustomer = GetCustomer(order, debug);

                if (DBCustomer.Success && DBCustomer.customer != null)
                {
                    if(debug)
                        log?.LogInformation($"UpdateOrCreateDbOrder: Found customer for order {order.ExternalId} in CreateOrUpdateDB.");

                    order.Customer = DBCustomer.customer;
                }

                //try to get existing order from database
                DBOrder = services?.GetOrderFromDB(order.ExternalId, debug);

                if (DBOrder != null && DBOrder != default(Order))
                {
                    if(debug)
                        log?.LogInformation($"UpdateOrCreateDbOrder: Found existing order {order.ExternalId} in CreateOrUpdateDB.");

                    DBOrder = order;
                    DBOrder.Status = "Updated";
                    returnValue = DBOrder;
                    UpdateOrder(DBOrder, "order message info", debug);
                }
                else if(debug)
                {
                    log?.LogInformation($"UpdateOrCreateDbOrder: Order does not exist in database.");
                }

                if (returnValue == null)
                {
                    if(debug)
                        log?.LogInformation($"UpdateOrCreateDbOrder: Creating new order.");

                    Order NewOrder = new Order();

                    if (DBCustomer.Success && DBCustomer.customer != null)
                    {
                        if(debug)
                            log?.LogInformation($"UpdateOrCreateDbOrder: Setting new order customer.");

                        NewOrder.Customer = order.Customer;
                        NewOrder.CustomerName = order.CustomerName;
                        NewOrder.CustomerType = order.CustomerType;
                        NewOrder.CustomerNo = order.CustomerNo;
                        NewOrder.CustomerID = DBCustomer.customer.ID;
                    }

                    NewOrder.ID = Guid.NewGuid();
                    NewOrder.AdditionalInfo = order.AdditionalInfo;
                    NewOrder.Created = DateTime.Now;
                    NewOrder.ProjectManager = order.ProjectManager;
                    NewOrder.Seller = order.Seller;
                    NewOrder.No = order.ExternalId;
                    NewOrder.ExternalId = order.ExternalId;
                    NewOrder.Status = "New";
                    NewOrder.Type = order.Type;
                    NewOrder.Handled = false;

                    try
                    {
                        if(debug)
                            log?.LogInformation($"UpdateOrCreateDbOrder: Adding order to DB.");

                        if (services?.AddOrderInDB(NewOrder, debug) == true)
                        {
                            if(debug)
                                log?.LogInformation($"UpdateOrCreateDbOrder: Fetching added order from DB.");

                            Order? NewDBOrder = services?.GetOrderFromDB(NewOrder.ExternalId, debug);

                            if (NewDBOrder != null)
                            {
                                returnValue = NewDBOrder;
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        log?.LogError("UpdateOrCreateDbOrder: " + ex.ToString());

                        if (debug)
                        {
                            log?.LogInformation($"UpdateOrCreateDbOrder: Failed to create order in database with error: {ex}");
                        }
                    }
                }
            }

            return returnValue;
        }

        public Customer? UpdateOrCreateDbCustomer(CustomerMessage? msg, bool debug)
        {
            Customer? returnValue = null;
            List<Customer> DBCustomers = new List<Customer>();

            if(msg == null || services == null)
            {
                return null;
            }

            try
            {
                if (debug)
                    log?.LogInformation($"UpdateOrCreateDbCustomer: Try to get customer {msg.CustomerNo} ({msg.Type}) from database");

                DBCustomers = services.GetCustomerFromDB(msg.CustomerNo, msg.Type, debug);

                if (DBCustomers.Count > 0)
                {
                    //If existing customer found, update it
                    Customer? foundCustomer = DBCustomers.FirstOrDefault(c => c.ExternalId == msg.CustomerNo && c.Type == msg.Type);

                    if (foundCustomer != null && foundCustomer != default(Customer))
                    {
                        if (debug)
                            log?.LogInformation($"UpdateOrCreateDbCustomer: Found customer {msg.CustomerNo} ({msg.Type}) in database so we update it.");

                        foundCustomer.Address = msg.CustomerAddress;
                        foundCustomer.Address1 = msg.CustomerAddress2;
                        foundCustomer.Fax = msg.CustomerFax;
                        foundCustomer.Phone = msg.CustomerPhone;
                        foundCustomer.City = msg.CustomerCity;
                        foundCustomer.Country = msg.CustomerCountry;
                        foundCustomer.ZipCode = msg.CustomerZipCode;
                        foundCustomer.State = msg.CustomerState;
                        foundCustomer.ProjectManager = msg.Responsible;
                        foundCustomer.Seller = msg.Responsible;
                        foundCustomer.Prospect = msg.Responsible;
                        foundCustomer.Modified = DateTime.Now;

                        if (debug)
                            log?.LogInformation($"UpdateOrCreateDbCustomer: Update customer {msg.CustomerNo} ({msg.Type}) with information: " + JsonConvert.SerializeObject(foundCustomer));

                        UpdateCustomer(foundCustomer, "new customer info", debug);
                        returnValue = foundCustomer;
                    }
                }
            }
            catch (Exception ex)
            {
                log?.LogError("UpdateOrCreateDbCustomer: " + ex.ToString());

                if (debug)
                {
                    log?.LogInformation($"UpdateOrCreateDbCustomer: Failed to find or update customer {msg.CustomerName} ({msg.CustomerNo}) in database with error: {ex}");
                }
            }

            if (DBCustomers.Count <= 0)
            {
                //No matching customers in db, create a new record
                Customer newCustomer = new Customer();
                newCustomer.Address = msg.CustomerAddress;
                newCustomer.Address1 = msg.CustomerAddress2;
                newCustomer.Fax = msg.CustomerFax;
                newCustomer.Phone = msg.CustomerPhone;
                newCustomer.City = msg.CustomerCity;
                newCustomer.Country = msg.CustomerCountry;
                newCustomer.ZipCode = msg.CustomerZipCode;
                newCustomer.State = msg.CustomerState;
                newCustomer.ProjectManager = msg.Responsible;
                newCustomer.Seller = msg.Responsible;
                newCustomer.Prospect = msg.Responsible;
                newCustomer.Name = msg.CustomerName;
                newCustomer.Type = msg.Type;
                newCustomer.ExternalId = msg.CustomerNo;
                newCustomer.ID = Guid.NewGuid();
                newCustomer.Created = DateTime.Now;
                newCustomer.Modified = DateTime.Now;

                try
                {
                    if (debug)
                        log?.LogInformation($"UpdateOrCreateDbCustomer: Trying to creating customer {msg.CustomerNo} ({msg.Type}) with information: " + JsonConvert.SerializeObject(newCustomer));

                    if (services.AddCustomerInDB(newCustomer, debug))
                    {
                        DBCustomers = services.GetCustomerFromDB(msg.CustomerNo, msg.Type, debug);

                        if (DBCustomers.Count > 0)
                        {
                            //If existing customer found, update it
                            Customer? foundCustomer = DBCustomers.FirstOrDefault(c => c.ExternalId == msg.CustomerNo && c.Type == msg.Type && c.Name == msg.CustomerName);

                            if (foundCustomer != null && foundCustomer != default(Customer))
                            {
                                returnValue = foundCustomer;
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    if (debug)
                    {
                        log?.LogError(ex.ToString());
                        log?.LogTrace($"Failed to add customer {newCustomer.Name} in database with error: {ex}");
                    }
                }
            }

            return returnValue;
        }

        public FindCustomerResult GetCustomer(Order? order, bool debug)
        {
            FindCustomerResult returnValue = new FindCustomerResult();

            if (order?.Customer == null)
            {
                if(debug)
                    log?.LogInformation($"GetCustomer: Order object customer was not set so fecth customer from database");

                FindCustomerResult foundCustomers = GetCustomer(order?.CustomerNo, order?.CustomerType, debug);

                if (string.IsNullOrEmpty(order?.CustomerName))
                {
                    if (foundCustomers.Success && foundCustomers.customers.Count > 0)
                    {
                        Customer? DBCustomer = foundCustomers.customers.OrderByDescending(c => c.Created).Take(1).FirstOrDefault();

                        if (DBCustomer != null && DBCustomer != default(Customer))
                        {
                            if (debug)
                                log?.LogInformation($"GetCustomer: Found customer {DBCustomer.ID} in database.");

                            returnValue.Success = true;
                            returnValue.customer = DBCustomer;
                        }
                    }
                }
                else
                {
                    if (foundCustomers.Success && foundCustomers.customers.Count > 0)
                    {
                        Customer? DBCustomer = foundCustomers.customers.FirstOrDefault(c => c.Name == order.CustomerName);

                        if (DBCustomer != null && DBCustomer != default(Customer))
                        {
                            if (debug)
                                log?.LogInformation($"GetCustomer: Found customer {DBCustomer.ID} in database.");

                            returnValue.Success = true;
                            returnValue.customer = DBCustomer;
                        }
                    }
                }
            }
            else
            {
                if (debug)
                    log?.LogInformation($"GetCustomer: Order has customer {order.Customer.ID}.");

                returnValue.Success = true;
                returnValue.customer = order.Customer;
            }

            return returnValue;
        }

        public FindCustomerResult GetCustomer(string CustomerNo, string CustomerType, string CustomerName, bool debug)
        {
            FindCustomerResult returnValue = new FindCustomerResult();

            if (debug)
                log?.LogInformation($"GetCustomer: Try to find customer {CustomerNo} ({CustomerType}) in database");

            FindCustomerResult foundCustomers = GetCustomer(CustomerNo, CustomerType, debug);

            if (foundCustomers.Success && foundCustomers.customers.Count > 0 && !string.IsNullOrEmpty(CustomerName))
            {
                Customer? DBCustomer = foundCustomers.customers.FirstOrDefault(c => c.ExternalId == CustomerNo && c.Type == CustomerType);

                if (DBCustomer != null && DBCustomer != default(Customer))
                {
                    if (debug)
                        log?.LogInformation($"GetCustomer: Found customer {DBCustomer.ID} in database.");

                    returnValue.Success = true;
                    returnValue.customer = DBCustomer;
                }
            }
            else if(foundCustomers.Success && foundCustomers.customers.Count > 0)
            {
                Customer? DBCustomer = foundCustomers.customers.OrderByDescending(c => c.Created).Take(1).FirstOrDefault();

                if (DBCustomer != null && DBCustomer != default(Customer))
                {
                    if (debug)
                        log?.LogInformation($"GetCustomer: Found customer {DBCustomer.ID} in database.");

                    returnValue.Success = true;
                    returnValue.customer = DBCustomer;
                }
            }

            return returnValue;
        }

        public FindCustomerResult GetCustomer(string? CustomerNo, string? CustomerType, bool debug)
        {
            FindCustomerResult returnValue = new FindCustomerResult();
            returnValue.Success = false;
            List<Customer>? dbCustomer = new List<Customer>();

            if (string.IsNullOrEmpty(CustomerNo))
            {
                if(debug)
                    log?.LogError($"GetCustomer: CustomerNo is null or empty");

                return returnValue;
            }

            try
            {
                if (debug)
                    log?.LogInformation($"GetCustomer: Trying to get customer {CustomerNo} ({CustomerType}) from database.");

                dbCustomer = services?.GetCustomerFromDB(CustomerNo, CustomerType, debug);

                if (dbCustomer?.Count > 0)
                {
                    if(debug)
                        log?.LogInformation($"GetCustomer: Found {dbCustomer.Count} customers. Returning first match if count was more than 1.");

                    returnValue.customer = dbCustomer[0];
                    returnValue.customers = dbCustomer;
                    returnValue.Success = true;
                }
            }
            catch (Exception ex)
            {
                log?.LogError("GetCustomer: " + ex.ToString());

                if (debug)
                {
                    log?.LogInformation($"GetCustomer: Failed to get customer {CustomerNo} from DB with error: " + ex.ToString());
                }
            }

            return returnValue;
        }

        public void UpdateCustomer(Customer customer, string logmsg, bool debug)
        {
            if (customer.ID == Guid.Empty)
            {
                if(debug)
                    log?.LogError($"UpdateCustomer: Failed to update {logmsg} on customer {customer.Name} ({customer.ExternalId}): No database record exists");

                return;
            }

            string updatequery = "";

            try
            {
                customer.Modified = DateTime.Now;

                if (debug)
                {
                    if (services?.UpdateCustomerInDB(customer) == true)
                    {
                        log?.LogInformation($"UpdateCustomer: Updated {logmsg} for customer {customer.Name} ({customer.ExternalId}) in database.");
                    }
                    else
                    {
                        log?.LogTrace($"UpdateCustomer: Failed to update {logmsg} on customer {customer.Name} ({customer.ExternalId})");
                    }
                }
            }
            catch (Exception ex)
            {
                log?.LogError("UpdateCustomer: " + ex.ToString());

                if (debug)
                {
                    log?.LogInformation($"UpdateCustomer: Failed to update {logmsg} on customer {customer.Name} ({customer.ExternalId}) with query: {updatequery}");
                }
            }
        }

        public string GetMailNickname(string? customerName, string? customerNo, string? customerType, bool debug)
        {
            string mailNickname = "";

            if(!string.IsNullOrEmpty(customerName) && !string.IsNullOrEmpty(customerNo))
            {
                if (customerType == "Customer")
                    mailNickname = RE.Regex.Replace(customerName, @"[^\w-]", "", RE.RegexOptions.None, TimeSpan.FromSeconds(1.5)) + "-" + customerNo + "-Kund";
                if (customerType == "Supplier")
                    mailNickname = RE.Regex.Replace(customerName, @"[^\w-]", "", RE.RegexOptions.None, TimeSpan.FromSeconds(1.5)) + "-" + customerNo + "-Lev";
            }

            illegalChars.ForEach(c => mailNickname = mailNickname.Replace(c.ToString(), ""));

            if (debug)
                log?.LogInformation($"GetMailNickname: {mailNickname.Replace("é", "e")}");

            return mailNickname.Replace("é", "e");
        }

        public FindCustomerGroupResult FindCustomerGroupAndDrive(string? customerName, string? customerNo, string? customerType, bool debug)
        {
            FindCustomerGroupResult returnValue = new FindCustomerGroupResult();
            returnValue.Success = false;
            string mailNickname = this.GetMailNickname(customerName, customerNo, customerType, debug);

            if(string.IsNullOrEmpty(customerName) || string.IsNullOrEmpty(customerNo) || string.IsNullOrEmpty(customerType))
            {
                return returnValue;
            }

            if(debug)
                log?.LogInformation($"FindCustomerGroupAndDrive: Trying to get group for {mailNickname}.");

            try
            {
                FindCustomerResult findCustomer = GetCustomer(customerNo, customerType, customerName, debug);

                if (findCustomer.Success)
                {
                    returnValue.customer = findCustomer.customer;

                    if (returnValue.customer != null)
                    {
                        returnValue = FindCustomerGroupAndDrive(returnValue.customer, debug).Result;
                    }
                }
            }
            catch (Exception ex)
            {
                log?.LogError(ex.ToString());

                if(debug)
                    log?.LogInformation($"FindCustomerGroupAndDrive: Failed to get group and drive for {mailNickname} with error: " + ex.ToString());
            }

            return returnValue;
        }

        public async Task<FindCustomerGroupResult> FindCustomerGroupAndDrive(Customer? customer, bool debug)
        {
            FindCustomerGroupResult returnValue = new FindCustomerGroupResult();
            returnValue.Success = false;
            string? groupDriveId = "";
            List<DriveItem> rootItems = new List<DriveItem>();
            List<DriveItem> generalItems = new List<DriveItem>();
            FindGroupResult result = new FindGroupResult() { Success = false };

            if(msGraph == null || customer == null)
            {
                return returnValue;
            }

            if(debug)
                log?.LogTrace($"FindCustomerGroupAndDrive: Trying to get group drive for {customer.Name}.");

            if (!string.IsNullOrEmpty(customer.GroupID))
            {
                result = await msGraph.GetGroupById(customer.GroupID, debug);
            }
            else
            {
                if(debug)
                    log?.LogTrace($"FindCustomerGroupAndDrive: Missing group id in database so trying mailnickname {customer.Name}.");

                string mailNickname = this.GetMailNickname(customer.Name, customer.ExternalId, customer.Type, debug);
                result = await msGraph.FindGroupByName(mailNickname, false);
            }

            if (result.Success)
            {
                returnValue.customer = customer;

                if (result.Count > 1 && result.groups != null)
                {
                    if(debug)
                        log?.LogTrace($"FindCustomerGroupAndDrive: Found multiple groups for {customer.Name}. Returning first match.");

                    returnValue.groupId = result.groups[0];
                    groupDriveId = await msGraph.GetGroupDrive(result.groups[0], debug);
                }
                else
                {
                    if(debug)
                        log?.LogTrace($"FindCustomerGroupAndDrive: Found group for {customer.Name}.");

                    returnValue.groupId = result.group;
                    groupDriveId = await msGraph.GetGroupDrive(result.group, debug);
                }

                if(!string.IsNullOrEmpty(returnValue.groupId))
                {
                    returnValue.customer.GroupID = returnValue.groupId ?? "";
                    returnValue.customer.GroupCreated = true;
                    UpdateCustomer(returnValue.customer, "group info", debug);
                }

                if (!string.IsNullOrEmpty(groupDriveId))
                {
                    if(debug)
                        log?.LogInformation($"FindCustomerGroupAndDrive: Found group drive for {customer.Name}.");

                    returnValue.Success = true;
                    returnValue.groupDriveId = groupDriveId;
                    returnValue.customer.DriveID = groupDriveId ?? "";
                    UpdateCustomer(returnValue.customer, "drive info", debug);
                    rootItems = await msGraph.GetDriveRootItems(groupDriveId, debug);

                    if (rootItems.Count > 0)
                    {
                        if(debug)
                            log?.LogInformation($"FindCustomerGroupAndDrive: Fetched root items in group drive for {customer.Name}.");

                        returnValue.rootItems = rootItems;
                    }
                }

                if (rootItems.Count > 0)
                {
                    var generalFolder = rootItems.FirstOrDefault(ri => ri.Name == "General");

                    if (generalFolder != default(DriveItem))
                    {
                        if (debug)
                            log?.LogInformation($"FindCustomerGroupAndDrive: Fetched general folder in group drive for {customer.Name}.");

                        returnValue.generalFolder = generalFolder;
                        returnValue.customer.GeneralFolderID = generalFolder.Id ?? "";
                        returnValue.customer.GeneralFolderCreated = true;
                        UpdateCustomer(returnValue.customer, "general folder info", debug);
                    }
                }
            }

            return returnValue;
        }

        public FindOrderGroupAndFolder GetOrderGroupAndFolder(string OrderNo, bool debug)
        {
            FindOrderGroupAndFolder returnValue = new FindOrderGroupAndFolder();
            returnValue.Success = false;
            Order? order = this.GetOrderFromCDN(OrderNo, debug);

            if(debug)
                log?.LogInformation($"GetOrderGroupAndFolder: Trying to fetch CDN item for {OrderNo}.");

            if(order != null)
            {
                returnValue = this.GetOrderGroupAndFolder(order, debug).Result;
            }

            return returnValue;
        }

        public async Task<FindOrderGroupAndFolder> GetOrderGroupAndFolder(Order order, bool debug)
        {
            FindOrderGroupAndFolder returnValue = new FindOrderGroupAndFolder();
            returnValue.Success = false;

            if(settings == null || settings.GraphClient == null || msGraph == null)
            {
                return returnValue;
            }

            if (order != null && !string.IsNullOrEmpty(order.CustomerNo) && !string.IsNullOrEmpty(order.CustomerType))
            {
                FindCustomerResult customerName = GetCustomer(order, debug);

                if (customerName.Success && customerName.customer != null)
                {
                    if(debug)
                        log?.LogInformation($"GetOrderGroupAndFolder: Got customer name from cdn for {customerName.customer.Name}.");

                    order.Customer = customerName.customer;
                    returnValue.customer = customerName.customer;

                    FindCustomerGroupResult? findCustomerGroupResult = new FindCustomerGroupResult();

                    if (returnValue.customer != null)
                    {
                        if(string.IsNullOrEmpty(returnValue.customer.GroupID) || string.IsNullOrEmpty(returnValue.customer.DriveID))
                        {
                            findCustomerGroupResult = this.FindCustomerGroupAndDrive(returnValue.customer.Name, returnValue.customer.ExternalId, returnValue.customer.Type, debug);
                        }
                        else
                        {
                            findCustomerGroupResult = await this.FindCustomerGroupAndDrive(returnValue.customer, debug);
                        }
                    }

                    if (findCustomerGroupResult?.Success == true && 
                        returnValue.customer != null && 
                        findCustomerGroupResult?.groupId != null && 
                        findCustomerGroupResult?.groupDriveId != null)
                    {
                        returnValue.customer.GroupID = findCustomerGroupResult.groupId ?? "";
                        returnValue.customer.DriveID = findCustomerGroupResult.groupDriveId ?? "";
                        returnValue.customer.GroupCreated = true;
                        UpdateCustomer(returnValue.customer, "group and drive info", debug);
                        
                        if(debug)
                            log?.LogInformation($"GetOrderGroupAndFolder: Found group for {returnValue.customer.Name} and order {order.ExternalId}.");

                        try
                        {
                            returnValue.orderTeamId = await msGraph.GetTeamFromGroup(returnValue.customer.GroupID, debug);

                            if (returnValue.orderTeamId != null)
                            {
                                returnValue.customer.TeamCreated = true;
                                returnValue.customer.TeamID = returnValue.orderTeamId ?? "";
                            }

                            UpdateCustomer(returnValue.customer, "team info", debug);

                            if(debug)
                                log?.LogInformation($"GetOrderGroupAndFolder: Found team for {returnValue.customer.Name} and order {order.ExternalId}.");
                        }
                        catch (Exception ex)
                        {
                            log?.LogError("GetOrderGroupAndFolder: " + ex.ToString());

                            if(debug)
                                log?.LogInformation($"GetOrderGroupAndFolder: Failed to find team for {returnValue.customer.Name} and order {order.ExternalId}.");
                        }

                        returnValue.Success = true;
                        returnValue.orderGroupId = findCustomerGroupResult.groupId;
                        returnValue.orderDriveId = findCustomerGroupResult.groupDriveId;

                        if (findCustomerGroupResult.generalFolder != null)
                        {
                            if(debug)
                                log?.LogInformation($"GetOrderGroupAndFolder: Found general folder for {returnValue.customer.Name} and order {order.ExternalId}.");

                            returnValue.generalFolder = findCustomerGroupResult.generalFolder;
                            returnValue.customer.GeneralFolderCreated = true;
                            returnValue.customer.GeneralFolderID = returnValue.generalFolder.Id ?? "";
                            UpdateCustomer(returnValue.customer, "general folder info", debug);
                        }

                        if (returnValue.customer.GeneralFolderCreated && msGraph != null)
                        {
                            string parentName = "";

                            switch (order.Type)
                            {
                                case "Order":
                                    parentName = "Order";
                                    break;
                                case "Project":
                                    parentName = "Order";
                                    break;
                                case "Quote":
                                    parentName = "Offert";
                                    RE.Match orderMatch = RE.Regex.Match(order.ExternalId, @"^([A-Z]?\d+)");

                                    if (orderMatch.Success)
                                    {
                                        if(debug)
                                            log?.LogInformation($"GetOrderGroupAndFolder: Changed order no for quote: {order.ExternalId} to: {orderMatch.Value}");

                                        order.ExternalId = orderMatch.Value;
                                    }

                                    break;
                                case "Offer":
                                    parentName = "Offert";
                                    RE.Match offerMatch = RE.Regex.Match(order.ExternalId, @"^([A-Z]?\d+)");

                                    if (offerMatch.Success)
                                    {
                                        if(debug)
                                            log?.LogInformation($"Changed order no for quote: {order.ExternalId} to: {offerMatch.Value}");

                                        order.ExternalId = offerMatch.Value;
                                    }

                                    break;
                                case "Purchase":
                                    parentName = "Beställning";
                                    break;
                                default:
                                    break;
                            }

                            try
                            {
                                DriveItem? foundOrderFolder = await msGraph.FindItem(returnValue.orderDriveId, "General/" + parentName + "/" + order.ExternalId, false);

                                if (foundOrderFolder != null)
                                {
                                    if(debug)
                                        log?.LogInformation($"GetOrderGroupAndFolder: Found order folder for {order.ExternalId} in customer/supplier {returnValue.customer.Name}.");

                                    returnValue.orderFolder = foundOrderFolder;
                                    order.CreatedFolder = true;
                                    order.CustomerID = returnValue.customer.ID;
                                    order.GroupFound = true;
                                    order.GeneralFolderFound = true;
                                    order.FolderID = returnValue.orderFolder.Id ?? "";
                                    order.OrdersFolderFound = true;
                                    UpdateOrder(order, "folder info", debug);
                                }
                                else
                                {
                                    List<DriveItem> rootItems = await msGraph.GetDriveRootItems(returnValue.orderDriveId, debug);

                                    foreach(DriveItem rootItem in rootItems)
                                    {
                                        if(rootItem.Name == "General")
                                        {
                                            List<DriveItem> generalItems = await msGraph.GetDriveFolderChildren(returnValue.orderDriveId, rootItem.Id, false, debug);

                                            foreach (DriveItem generalItem in generalItems)
                                            {
                                                if (generalItem.Name == parentName)
                                                {
                                                    List<DriveItem> folderItems = await msGraph.GetDriveFolderChildren(returnValue.orderDriveId, generalItem.Id, false, debug);

                                                    foreach(DriveItem folderItem in folderItems)
                                                    {
                                                        if(folderItem.Name == order.ExternalId)
                                                        {
                                                            if(debug)
                                                                log?.LogInformation($"GetOrderGroupAndFolder: Found order folder for {order.ExternalId} in customer/supplier {returnValue.customer.Name}.");

                                                            returnValue.orderFolder = folderItem;
                                                            order.CreatedFolder = true;
                                                            order.CustomerID = returnValue.customer.ID;
                                                            order.GroupFound = true;
                                                            order.GeneralFolderFound = true;
                                                            order.FolderID = returnValue.orderFolder.Id ?? "";
                                                            order.OrdersFolderFound = true;
                                                            UpdateOrder(order, "folder info", debug);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                log?.LogError("GetOrderGroupAndFolder: " + ex.ToString());

                                if(debug)
                                    log?.LogInformation($"Failed to get folder for order {order.ExternalId}.");
                            }
                        }
                    }
                }
            }

            return returnValue;
        }

        public string GetOrderParentFolderName(string orderType)
        {
            string parentName = "";

            switch (orderType)
            {
                case "Order":
                    parentName = "Order";
                    break;
                case "Project":
                    parentName = "Order";
                    break;
                case "Quote":
                    parentName = "Offert";
                    break;
                case "Offer":
                    parentName = "Offert";
                    break;
                case "Purchase":
                    parentName = "Beställning";
                    break;
                default:
                    break;
            }

            return parentName;
        }

        public string GetOrderExternalId(string orderType, string orderNo)
        {
            string returnValue = orderNo;

            if (orderType == "Quote" || orderType == "Offer")
            {
                RE.Match offerMatch = RE.Regex.Match(orderNo, quoteRegex);

                if (offerMatch.Success)
                {
                    returnValue = offerMatch.Value;
                }
            }

            return returnValue;
        }

        public string FindOrderNoInString(string input)
        {
            string returnValue = "";
            RE.Match orderMatches = RE.Regex.Match(input, orderRegex);

            if (orderMatches.Success)
            {
                returnValue = orderMatches.Value;
            }

            return returnValue;
        }

        public string FindCustomerNoInString(string input)
        {
            string returnValue = "";
            RE.Match orderMatches = RE.Regex.Match(input, customerRegex);

            if (orderMatches.Success)
            {
                returnValue = orderMatches.Groups[1].Value;
            }

            return returnValue;
        }

        public List<Order> GetUnhandledOrderItems(bool debug)
        {
            List<Order> returnValue = new List<Order>();

            if (services == null)
            {
                return returnValue;
            }

            returnValue = services.ExecSQLQuery<Order>("SELECT * FROM Orders WHERE Handled = 0", new Dictionary<string, object>(), debug);
            return returnValue;
        }

        public async Task<DriveItem?> GetEmailsFolder(string parent, string month, string year, bool debug)
        {
            string? groupDriveId = "";

            if (settings == null || settings.GraphClient == null || msGraph == null || string.IsNullOrEmpty(CDNTeamID))
            {
                return null;
            }

            try
            {
                groupDriveId = await msGraph.GetGroupDrive(CDNTeamID, debug);
            }
            catch (Exception ex)
            {
                log?.LogError("GetEmailsFolder: " + ex.ToString());

                if(debug)
                    log?.LogTrace($"GetEmailsFolder: Failed to get drive for CDN Team with error: " + ex.ToString());
            }

            DriveItem? emailFolder = default(DriveItem);

            if (!string.IsNullOrEmpty(groupDriveId))
            {
                try
                {
                    emailFolder = await msGraph.FindItem(groupDriveId, parent + "/EmailMessages_" + month + "_" + year, false);
                }
                catch (Exception ex)
                {
                    log?.LogError("GetEmailsFolder: " + ex.ToString());

                    if(debug)
                        log?.LogInformation($"GetEmailsFolder: Failed to get email folder for CDN Team with error: " + ex.ToString());
                }
            }

            return emailFolder;
        }

        public async Task<DriveItem?> GetGeneralFolder(string groupId, bool debug)
        {
            string? groupDriveId = "";

            if (settings == null || settings.GraphClient == null || msGraph == null)
            {
                return null;
            }

            try
            {
                groupDriveId = await msGraph.GetGroupDrive(groupId, debug);
            }
            catch (Exception ex)
            {
                log?.LogError("GetGeneralFolder: " + ex.ToString());

                if(debug)
                    log?.LogInformation($"GetGeneralFolder: Failed to get drive for group {groupId} with error: " + ex.ToString());
            }

            DriveItem? generalFolder = default(DriveItem);

            if (!string.IsNullOrEmpty(groupDriveId))
            {
                try
                {
                    generalFolder = await msGraph.FindItem(groupDriveId, "General", false);
                }
                catch (Exception ex)
                {
                    log?.LogError("GetGeneralFolder: " + ex.ToString());
                    log?.LogTrace($"GetGeneralFolder: Failed to get general folder in group {groupId} with error: " + ex.ToString());
                }
            }

            return generalFolder;
        }

        public async Task<DriveItem?> GetOrderFolder(string groupId, string groupDriveId, Order order, bool debug)
        {
            DriveItem? returnValue = null;

            if(order.Type == null || string.IsNullOrEmpty(groupId) || msGraph == null)
            {
                return null;
            }

            string parentName = GetOrderParentFolderName(order.Type);
            DriveItem? generalFolder = await GetGeneralFolder(groupId, debug);

            if(generalFolder != null)
            {
                DriveItem? orderParentFolder = await msGraph.FindItem(groupDriveId, generalFolder.Id, parentName, false, debug);
                
                if(orderParentFolder != null)
                {
                    var orderfolderName = GetOrderExternalId(order.Type, order.ExternalId);
                    DriveItem? orderFolder = await msGraph.FindItem(groupDriveId, orderParentFolder.Id, orderfolderName, false, debug);

                    if(orderFolder != null)
                    {
                        returnValue = orderFolder;
                    }
                }
            }

            return returnValue;
        }

        public Order SetFolderStatus(Order order, bool found)
        {
            Order returnValue = order;

            if (order.Type == "Order" || order.Type == "Project")
            {
                returnValue.OffersFolderFound = false;
                returnValue.PurchaseFolderFound = false;
                returnValue.OrdersFolderFound = found;
            }
            else if (order.Type == "Quote" || order.Type == "Offer")
            {
                returnValue.OffersFolderFound = found;
                returnValue.PurchaseFolderFound = false;
                returnValue.OrdersFolderFound = false;
            }
            else if (order.Type == "Purchase")
            {
                returnValue.OffersFolderFound = false;
                returnValue.PurchaseFolderFound = found;
                returnValue.OrdersFolderFound = false;
            }

            return returnValue;
        }

        public async Task<List<DriveItem>> GetOrderTemplateFolders(Order order, bool debug)
        {
            List<DriveItem> foldersToCreate = new List<DriveItem>();

            if(msGraph == null || settings == null)
            {
                return foldersToCreate;
            }

            var cdnDrive = await msGraph.GetSiteDrive(settings.cdnSiteId, debug);

            if (cdnDrive != null)
            {
                DriveItem? folder = await msGraph.FindItem(cdnDrive, "Dokumentstruktur " + order.Type, false);

                if(folder != null)
                {
                    List<DriveItem> folderChildren = await msGraph.GetDriveFolderChildren(cdnDrive, folder.Id, true, debug);
                    foldersToCreate.AddRange(folderChildren);
                }
            }

            return foldersToCreate;
        }

        public async Task<CreateCustomerResult> CreateCustomerGroup(Customer customer, bool debug)
        {
            CreateCustomerResult returnValue = new CreateCustomerResult();
            returnValue.Success = false;
            string? group = "";

            if (settings == null || settings.GraphClient == null || msGraph == null || customer == null)
            {
                return returnValue;
            }

            string[]? admins = null;
            
            if (!string.IsNullOrEmpty(settings.Admins))
            {
                admins = settings?.Admins.Split(',');
            }

            List<string> adminids = new List<string>();
            string mailNickname = this.GetMailNickname(customer.Name, customer.ExternalId, customer.Type, debug);
            adminids = await GetAdmins(new Customer(), admins, debug);
            string GroupName = "";

            if (customer.Type == "Customer")
                GroupName = customer.Name + " (" + customer.ExternalId + ") - Kund";
            if (customer.Type == "Supplier")
                GroupName = customer.Name + " (" + customer.ExternalId + ") - Lev";

            try
            {
                //Create a group without owners
                group = (await msGraph.CreateGroup(GroupName, mailNickname, adminids, debug))?.Id;

                if(debug)
                    log?.LogInformation($"CreateCustomerGroup: Created group for customer {customer.Name} ({customer.ExternalId})");
            }
            catch (Exception ex)
            {
                log?.LogError("CreateCustomerGroup: " + ex.ToString());

                if(debug)
                    log?.LogInformation($"CreateCustomerGroup: Failed to create group for {customer.Name} ({customer.ExternalId}) with error: " + ex.ToString());
            }

            //if the group was created
            if(group != null)
            {
                customer.GroupID = group ?? "";

                //get the group drive (will probably fail since thr group takes a while to create)
                try
                {
                    string? groupDriveId = await msGraph.GetGroupDrive(group, debug);

                    if (groupDriveId != null)
                    {
                        customer.DriveID = groupDriveId ?? "";
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("CreateCustomerGroup: " + ex.ToString());
                }

                returnValue.group = group;
                returnValue.customer = customer;
                returnValue.Success = true;
            }
            else
            {
                log?.LogInformation($"CreateCustomerGroup: Failed to create group {customer.Name} ({customer.ExternalId})");
            }

            return returnValue;
        }

        public async Task<bool> CreateCustomerTeam(Customer customer, string groupId, bool debug)
        {
            bool returnValue = false;

            if (settings == null || settings.GraphClient == null || msGraph == null || customer == null || string.IsNullOrEmpty(groupId))
            {
                return returnValue;
            }

            var appId = settings.config["CustomerCardAppId"];

            //try to get team or create it if it's missing
            var team = await msGraph.CreateTeamFromGroup(groupId, debug);

            if(debug)
                log?.LogInformation($"CreateCustomerTeam: Created team for {customer.Name} ({customer.ExternalId})");

            if (team != null)
            {
                customer.TeamCreated = true;
                customer.TeamID = team.Id ?? "";
                customer.TeamUrl = team.WebUrl ?? "";
                UpdateCustomer(customer, "team info", debug);

                try
                {
                    string ContentUrl = "https://holtabcustomercard.azurewebsites.net/Home/Index?id=" + team.Id;
                    string? teamId = team?.Id;

                    if (!customer.InstalledApp && !string.IsNullOrEmpty(groupId) && !string.IsNullOrEmpty(appId))
                    {
                        //string? groupDriveId = await msGraph.GetGroupDrive(groupId);
                        string? rootUrl = await msGraph.GetGroupDriveUrl(groupId, debug);

                        if (!string.IsNullOrEmpty(rootUrl) && !string.IsNullOrEmpty(teamId))
                        {
                            var channelSwedish = await msGraph.FindChannel(teamId, "Allmänt", debug);
                            var channelEnglish = await msGraph.FindChannel(teamId, "General", debug);
                            string? channel = channelSwedish ?? channelEnglish;

                            if (!string.IsNullOrEmpty(channel) && !string.IsNullOrEmpty(rootUrl))
                            {
                                var app = await msGraph.AddTeamApp(teamId, appId, debug);

                                if (app != null)
                                {
                                    if(debug)
                                        log?.LogInformation($"CreateCustomerTeam: Adding channel for app {app} to {customer.Name}");

                                    await msGraph.AddChannelApp(teamId, app, channel, "Om Företaget", System.Guid.NewGuid().ToString("D").ToUpperInvariant(), ContentUrl, rootUrl, "", debug);

                                    if(debug)
                                        log?.LogInformation($"CreateCustomerTeam: Installed teams app for {customer.Name} ({customer.ExternalId})");

                                    customer.InstalledApp = true;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("CreateCustomerTeam: " + ex.ToString());

                    if(debug)
                        log?.LogInformation($"CreateCustomerTeam: Failed to install teams app for {customer.Name} with error " + ex.ToString());
                }

                UpdateCustomer(customer, "team app info", debug);
                returnValue = true;
            }
            else
            {
                if(debug)
                    log?.LogInformation($"CreateCustomerTeam: Failed to create team for group {customer.Name} ({customer.ExternalId})");
            }

            return returnValue;
        }

        /// <summary>
        /// Create group and team for customer or supplier
        /// No checking is done if group or team already exists.
        /// </summary>
        /// <param name="customer"></param>
        /// <returns></returns>
        public async Task<CreateCustomerResult> CreateCustomerOrSupplier(Customer customer, bool debug)
        {
            CreateCustomerResult returnValue = new CreateCustomerResult();
            returnValue.Success = false;

            if(settings == null || settings.GraphClient == null || msGraph == null || customer == null)
            {
                return returnValue;
            }

            string? group = "";
            string[]? admins = null;

            if (!string.IsNullOrEmpty(settings.Admins))
            {
                admins = settings.Admins.Split(',');
            }

            List<string> adminids = new List<string>();
            string mailNickname = "";

            mailNickname = this.GetMailNickname(customer.Name, customer.ExternalId, customer.Type, debug);
            adminids = await GetAdmins(customer, admins, debug);
            string GroupName = "";

            if (customer.Type == "Customer")
                GroupName = customer.Name + " (" + customer.ExternalId + ") - Kund";
            if (customer.Type == "Supplier")
                GroupName = customer.Name + " (" + customer.ExternalId + ") - Lev";

            //find group if it exists or try to create it
            if (customer.GroupID != null && customer.GroupID != string.Empty)
            {
                FindGroupResult findGroup = await msGraph.GetGroupById(customer.GroupID, debug);

                if (findGroup?.Success == true)
                {
                    group = findGroup.group;
                }
            }
            else
            {
                //create group if it didn't exist
                try
                {
                    group = (await msGraph.CreateGroup(GroupName, mailNickname, adminids, debug))?.Id;

                    if(debug)
                        log?.LogInformation($"CreateCustomerOrSupplier: Created group for customer {customer.Name} ({customer.ExternalId})");
                }
                catch (Exception ex)
                {
                    log?.LogError(ex.ToString());

                    if(debug)
                        log?.LogInformation($"CreateCustomerOrSupplier: Failed to create group for {customer.Name} ({customer.ExternalId}) with error: " + ex.ToString());
                }
            }

            if (group != null)
            {
                customer.GroupID = group ?? "";

                try
                {
                    string? groupDriveId = await msGraph.GetGroupDrive(group, debug);

                    if (groupDriveId != null)
                    {
                        customer.DriveID = groupDriveId ?? "";
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("CreateCustomerOrSupplier: " + ex.ToString());
                }

                customer.GroupCreated = true;
                UpdateCustomer(customer, "group and drive info", debug);

                var team = await msGraph.CreateTeamFromGroup(customer.GroupID, debug);

                if(debug)
                    log?.LogInformation($"CreateCustomerOrSupplier: Created team for {customer.Name} ({customer.ExternalId})");

                if (team != null)
                {
                    customer.TeamCreated = true;
                    customer.TeamID = team.Id ?? "";
                    customer.TeamUrl = team.WebUrl ?? "";
                    UpdateCustomer(customer, "team info", debug);

                    try
                    {
                        string ContentUrl = "https://holtabcustomercard.azurewebsites.net/Home/Index?id=" + customer.TeamID;

                        if (!string.IsNullOrEmpty(group))
                        {
                            string? groupDriveId = await msGraph.GetGroupDrive(group, debug);
                            string? rootUrl = await msGraph.GetGroupDriveUrl(group, debug);

                            if (!string.IsNullOrEmpty(groupDriveId) && !string.IsNullOrEmpty(rootUrl))
                            {
                                var generalChannelSwedish = await msGraph.FindChannel(customer.TeamID, "Allmänt", debug);
                                var generalChannelEnglish = await msGraph.FindChannel(customer.TeamID, "General", debug);
                                string? generalChannel = generalChannelSwedish ?? generalChannelEnglish;

                                if (!string.IsNullOrEmpty(generalChannel))
                                {
                                    var app = await msGraph.AddTeamApp(customer.TeamID, "e2cb3981-47e7-47b3-a0e1-f9078d342253", debug);
                                        
                                    if(app != null)
                                    {
                                        await msGraph.AddChannelApp(customer.TeamID, app, generalChannel, "Om Företaget", System.Guid.NewGuid().ToString("D").ToUpperInvariant(), ContentUrl, rootUrl, "", debug);

                                        if(debug)
                                            log?.LogInformation($"CreateCustomerOrSupplier: Installed teams app for {customer.Name} ({customer.ExternalId})");

                                        customer.InstalledApp = true;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log?.LogError("CreateCustomerOrSupplier: " + ex.ToString());

                        if(debug)
                            log?.LogInformation($"CreateCustomerOrSupplier: Failed to install teams app for {customer.Name} with error: " + ex.ToString());
                    }

                    UpdateCustomer(customer, "team app info", debug);

                    returnValue.group = group;
                    returnValue.team = team;
                    returnValue.customer = customer;
                    returnValue.Success = true;
                }
                else
                {
                    if(debug)
                        log?.LogInformation($"CreateCustomerOrSupplier: Failed to create team for group {customer.Name} ({customer.ExternalId})");
                }
            }
            else
            {
                if(debug)
                    log?.LogInformation($"CreateCustomerOrSupplier: Failed to create group {customer.Name} ({customer.ExternalId})");
            }

            return returnValue;
        }

        /// <summary>
        /// Copy files and folders from structure in CDN site to group site for customer or supplier
        /// Expects the general folder to exist before the function is run
        /// </summary>
        /// <param name="customer"></param>
        /// <returns></returns>
        public async Task<bool> CopyRootStructure(Customer customer, bool debug)
        {
            bool returnValue = false;
            
            if(msGraph == null || string.IsNullOrEmpty(cdnSiteId) || customer == null)
            {
                return returnValue;
            }

            string? cdnDrive = await msGraph.GetSiteDrive(cdnSiteId, debug);

            if (!string.IsNullOrEmpty(cdnDrive))
            {
                DriveItem? source = default(DriveItem);

                try
                {
                    if (customer.Type == "Customer")
                    {
                        source = await msGraph.FindItem(cdnDrive, "Dokumentstruktur Kund", false, debug);
                    }
                    else if (customer.Type == "Supplier")
                    {
                        source = await msGraph.FindItem(cdnDrive, "Dokumentstruktur Leverantör", false, debug);
                    }
                }
                catch (Exception ex)
                {
                    log?.LogError("CopyRootStructure: " + ex.ToString());

                    if(debug)
                        log?.LogInformation($"CopyRootStructure: Failed to get templates for {customer.Name} with error " + ex.ToString());
                }

                if(debug)
                    log?.LogInformation($"CopyRootStructure: Found CDN folder structure template for {customer.Name} ({customer.ExternalId})");

                if (source != default(DriveItem))
                {
                    DriveItem? generalFolder = default(DriveItem);

                    if (string.IsNullOrEmpty(customer?.GeneralFolderID) && !string.IsNullOrEmpty(customer?.GroupID))
                    {
                        try
                        {
                            generalFolder = await this.GetGeneralFolder(customer.GroupID, debug);

                            if (generalFolder != null)
                            {
                                customer.GeneralFolderID = generalFolder.Id ?? "";

                                if(debug)
                                    log?.LogInformation($"CopyRootStructure: Found general folder for {customer.Name} ({customer.ExternalId})");
                            }
                        }
                        catch (Exception ex)
                        {
                            log?.LogError("CopyRootStructure: " + ex.ToString());
                            
                            if(debug)
                                log?.LogInformation($"CopyRootStructure: Failed to get general folder for {customer.Name} with error: " + ex.ToString());
                        }
                    }

                    if (!string.IsNullOrEmpty(customer?.GeneralFolderID) && !string.IsNullOrEmpty(customer?.GroupID))
                    {
                        try
                        {
                            var children = await msGraph.GetDriveFolderChildren(cdnDrive, source.Id, true, debug);

                            foreach (var child in children)
                            {
                                await msGraph.CopyFolder(customer.GroupID, customer.GeneralFolderID, child, true, false, debug);
                            }

                            if(debug)
                                log?.LogInformation($"CopyRootStructure: Copied templates for {customer.Name} ({customer.ExternalId})");

                            returnValue = true;
                        }
                        catch (Exception ex)
                        {
                            log?.LogError("CopyRootStructure: " + ex.ToString());

                            if(debug)
                                log?.LogInformation($"CopyRootStructure: Failed to copy template structure for {customer.Name} ({customer.ExternalId})");
                        }

                    }
                }
            }

            return returnValue;
        }

        public async Task<List<string>> GetAdmins(Customer? customer, string[]? admins, bool debug)
        {
            List<string> adminids = new List<string>();
            List<string> _admins = new List<string>();

            if (customer == null)
            {
                return adminids;
            }

            if(admins != null)
            {
                _admins.AddRange(admins);
            }

            //if seller exists add it to admins list
            if (!String.IsNullOrEmpty(customer?.Seller) && !_admins.Exists(a => a == customer.Seller))
                _admins.Add(customer.Seller);

            if(admins != null && settings != null && settings.GraphClient != null)
            {
                //Get all admin ids
                foreach (string user in admins)
                {
                    try
                    {
                        User? graphUser = await settings.GraphClient.Users[user].GetAsync();

                        if (graphUser != null)
                        {
                            adminids.Add("https://graph.microsoft.com/v1.0/users/" + graphUser.Id);
                        }
                        else
                        {
                            if(debug)
                                log?.LogInformation($"GetAdmins: Failed to find user {user}");
                        }
                    }
                    catch (Exception ex)
                    {
                        log?.LogError("GetAdmins: " + ex.ToString());

                        if(debug)
                            log?.LogInformation($"GetAdmins: Failed to get user {user}" + ex.ToString());
                    }
                }
            }

            return adminids;
        }
    }
}
