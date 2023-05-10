using CreateTeam.Models;
using Microsoft.ApplicationInsights;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using RE = System.Text.RegularExpressions;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Microsoft.ApplicationInsights.DataContracts;
using System.Linq;
using Microsoft.Data.SqlClient;
using PnP.Framework.Provisioning.Model.Teams;
using Microsoft.Graph.Models;

namespace CreateTeam.Shared
{
    public class Common
    {
        private readonly ILogger log;
        private readonly IConfiguration config;
        private readonly TelemetryClient telemetryClient;
        private readonly GraphServiceClient graphClient;
        private readonly Graph msGraph;
        private readonly string ClientID;
        private readonly string ClientSecret;
        private readonly string TenantID;
        private readonly string CDNTeamID;
        private readonly string cdnSiteId;
        private readonly string OrderListID;
        private readonly string CustomerListID;
        private readonly string SqlConnectionString;
        public readonly Group CDNGroup;
        public readonly Group EmailsFolder;
        public List<char> illegalChars = new List<char>() { '~', '`', '!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '_', '+', '=', '{', '}', '|', '[', ']', '\\', ':', '\"', ';', '\'', '<', '>', ',', '.', '?', '/', 'å', 'ä', 'ö', 'Å', 'Ä', 'Ö', ' ', 'Ø', 'Æ', 'æ', 'ø', 'ü', 'Ü', 'µ', 'ẞ', 'ß' };

        public Common(GraphServiceClient _graphClient, IConfiguration _config, ILogger _log, TelemetryClient _telemetryClient, Graph _msGraph)
        {
            log = _log;
            config = _config;
            graphClient = _graphClient;
            telemetryClient = _telemetryClient;
            msGraph = _msGraph;
            ClientID = _config["ClientID"];
            ClientSecret = _config["ClientSecret"];
            TenantID = _config["TenantID"];
            CDNTeamID = _config["CDNTeamID"];
            cdnSiteId = _config["cdnSiteId"];
            OrderListID = _config["OrderListID"];
            CustomerListID = _config["CustomerListID"];
            SqlConnectionString = _config["SqlConnectionString"];
        }

        public Order GetOrderFromCDN(string orderNo)
        {
            Order returnValue = new Order();

            try
            {
                returnValue = Services.GetOrderFromDB(orderNo, SqlConnectionString);
                FindCustomerResult customers = GetCustomer(returnValue);

                if (returnValue.Customer == null && customers.Success && customers.customer != null)
                {
                    returnValue.CustomerName = customers.customer.Name;
                    returnValue.Customer = customers.customer;

                    if (customers.customer.ID != Guid.Empty)
                    {
                        returnValue.CustomerID = customers.customer.ID;
                    }
                }

                returnValue.ExternalId = orderNo;
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace($"Failed to get order {orderNo} from CDN with error: " + ex.ToString());
            }

            UpdateOrder(returnValue, "customer info");

            return returnValue;
        }

        public void UpdateOrder(Order order, string logmsg)
        {
            if (order?.ID == Guid.Empty)
            {
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to update {logmsg} for order {order.ExternalId}: No database record exists"));
                return;
            }

            string updatequery = "";

            try
            {
                telemetryClient.TrackEvent(new EventTelemetry($"Updating {logmsg} for order {order.ExternalId} in database."));
                Dictionary<string, object> keys = new Dictionary<string, object>
                            {
                                { "ID", order.ID }
                            };

                updatequery = Services.GetUpdateSQLQuery(order, "Orders", keys);
                Services.UpdateOrderInDB(order, SqlConnectionString);
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to update {logmsg} for order {order.ExternalId} in database with query: {updatequery}"));
            }
        }

        public Order UpdateOrCreateDbOrder(Order order)
        {
            Order returnValue = null;
            Order DBOrder = null;

            //try to find the customer for the order
            FindCustomerResult DBCustomer = new FindCustomerResult();
            DBCustomer = GetCustomer(order);

            try
            {
                DBOrder = Services.GetOrderFromDB(order.ExternalId, SqlConnectionString);

                if (DBOrder != null && DBOrder != default(Order))
                {
                    DBOrder.AdditionalInfo = order.AdditionalInfo;
                    DBOrder.ProjectManager = order.ProjectManager;
                    DBOrder.Type = order.Type;
                    DBOrder.Seller = order.Seller;
                    DBOrder.Status = "Updated";

                    if (DBCustomer.Success && DBCustomer.customer != null)
                    {
                        DBOrder.Customer = order.Customer;
                        DBOrder.CustomerNo = order.CustomerNo;
                        DBOrder.CustomerType = order.CustomerType;
                        DBOrder.CustomerName = order.CustomerName;
                        DBOrder.CustomerID = DBCustomer.customer.ID;
                    }

                    returnValue = DBOrder;
                    UpdateOrder(DBOrder, "order message info");
                }
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to find or update order in database with error: {ex}"));
            }

            if (returnValue == null)
            {
                Order NewOrder = new Order();

                if (DBCustomer.Success && DBCustomer.customer != null)
                {
                    NewOrder.Customer = order.Customer;
                    NewOrder.CustomerName = order.CustomerName;
                    NewOrder.CustomerID = DBCustomer.customer.ID;
                }

                NewOrder.CustomerType = order.CustomerType;
                NewOrder.CustomerNo = order.CustomerNo;
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
                    Services.AddOrderInDB(NewOrder, SqlConnectionString);

                    Order NewDBOrder = Services.GetOrderFromDB(NewOrder.ExternalId, SqlConnectionString);

                    if (NewDBOrder != null)
                    {
                        returnValue = NewDBOrder;
                    }
                }
                catch (Exception ex)
                {
                    telemetryClient.TrackException(ex);
                    telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create order in database with error: {ex}"));
                }
            }

            return returnValue;
        }

        public Customer UpdateOrCreateDbCustomer(CustomerMessage msg)
        {
            Customer returnValue = null;
            List<Customer> DBCustomers = new List<Customer>();

            try
            {
                DBCustomers = Services.GetCustomerFromDB(msg.CustomerNo, msg.Type, SqlConnectionString);

                if (DBCustomers.Count > 0)
                {
                    //If existing customer found, update it
                    Customer foundCustomer = DBCustomers.FirstOrDefault(c => c.ExternalId == msg.CustomerNo && c.Type == msg.Type);

                    if (foundCustomer != null && foundCustomer != default(Customer))
                    {
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
                        UpdateCustomer(foundCustomer, "new customer info");
                        returnValue = foundCustomer;
                    }
                }
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to find or update customer {msg.CustomerName} ({msg.CustomerNo}) in database with error: {ex}"));
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
                    if (Services.AddCustomerInDB(newCustomer, SqlConnectionString))
                    {
                        DBCustomers = Services.GetCustomerFromDB(msg.CustomerNo, msg.Type, SqlConnectionString);

                        if (DBCustomers.Count > 0)
                        {
                            //If existing customer found, update it
                            Customer foundCustomer = DBCustomers.FirstOrDefault(c => c.ExternalId == msg.CustomerNo && c.Type == msg.Type && c.Name == msg.CustomerName);

                            if (foundCustomer != null && foundCustomer != default(Customer))
                            {
                                returnValue = foundCustomer;
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    telemetryClient.TrackException(ex);
                    telemetryClient.TrackTrace(new TraceTelemetry($"Failed to add customer {newCustomer.Name} in database with error: {ex}"));
                }
            }

            return returnValue;
        }

        public FindCustomerResult GetCustomer(Order order)
        {
            FindCustomerResult returnValue = new FindCustomerResult();

            if (order?.Customer == null)
            {
                FindCustomerResult foundCustomers = GetCustomer(order.CustomerNo, order.CustomerType);

                if (string.IsNullOrEmpty(order.CustomerName))
                {
                    if (foundCustomers.Success && foundCustomers.customers.Count > 0)
                    {
                        Customer DBCustomer = foundCustomers.customers.OrderByDescending(c => c.Created).Take(1).FirstOrDefault();

                        if (DBCustomer != null && DBCustomer != default(Customer))
                        {
                            returnValue.Success = true;
                            returnValue.customer = DBCustomer;
                        }
                    }
                }
                else
                {
                    if (foundCustomers.Success && foundCustomers.customers.Count > 0)
                    {
                        Customer DBCustomer = foundCustomers.customers.FirstOrDefault(c => c.Name == order.CustomerName);

                        if (DBCustomer != null && DBCustomer != default(Customer))
                        {
                            returnValue.Success = true;
                            returnValue.customer = DBCustomer;
                        }
                    }
                }
            }
            else
            {
                returnValue.Success = true;
                returnValue.customer = order.Customer;
            }

            return returnValue;
        }

        public FindCustomerResult GetCustomer(string CustomerNo, string CustomerType, string CustomerName)
        {
            FindCustomerResult returnValue = new FindCustomerResult();
            FindCustomerResult foundCustomers = GetCustomer(CustomerNo, CustomerType);

            if (foundCustomers.Success && foundCustomers.customers.Count > 0 && !string.IsNullOrEmpty(CustomerName))
            {
                Customer DBCustomer = foundCustomers.customers.FirstOrDefault(c => c.ExternalId == CustomerNo && c.Type == CustomerType);

                if (DBCustomer != null && DBCustomer != default(Customer))
                {
                    returnValue.Success = true;
                    returnValue.customer = DBCustomer;
                }
            }
            else if(foundCustomers.Success && foundCustomers.customers.Count > 0)
            {
                Customer DBCustomer = foundCustomers.customers.OrderByDescending(c => c.Created).Take(1).FirstOrDefault();

                if (DBCustomer != null && DBCustomer != default(Customer))
                {
                    returnValue.Success = true;
                    returnValue.customer = DBCustomer;
                }
            }

            return returnValue;
        }

        public FindCustomerResult GetCustomer(string CustomerNo, string CustomerType)
        {
            FindCustomerResult returnValue = new FindCustomerResult();
            returnValue.Success = false;
            List<Customer> dbCustomer = new List<Customer>();

            if (string.IsNullOrEmpty(CustomerNo))
            {
                return returnValue;
            }

            try
            {
                dbCustomer = Services.GetCustomerFromDB(CustomerNo, CustomerType, SqlConnectionString);

                if (dbCustomer.Count > 0)
                {
                    returnValue.Success = true;
                    returnValue.customers = dbCustomer;

                    if (dbCustomer.Count == 1)
                    {
                        returnValue.customer = dbCustomer[0];
                    }
                }
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace($"Failed to get customer {CustomerNo} from DB with error: " + ex.ToString());
            }

            return returnValue;
        }

        public void UpdateCustomer(Customer customer, string logmsg)
        {
            if (customer.ID == Guid.Empty)
            {
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to update {logmsg} on customer {customer.Name} ({customer.ExternalId}): No database record exists"));
                return;
            }

            string updatequery = "";

            try
            {
                customer.Modified = DateTime.Now;
                Dictionary<string, object> keys = new Dictionary<string, object>();
                keys.Add("ID", customer.ID);
                updatequery = Services.GetUpdateSQLQuery(customer, "Customers", keys);
                Services.UpdateCustomerInDB(customer, SqlConnectionString);
                telemetryClient.TrackEvent(new EventTelemetry($"Updated {logmsg} for customer {customer.Name} ({customer.ExternalId}) in database."));
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to update {logmsg} on customer {customer.Name} ({customer.ExternalId}) with query: {updatequery}"));
            }
        }

        public string GetMailNickname(string customerName, string customerNo, string customerType)
        {
            string mailNickname = "";

            if (customerType == "Customer")
                mailNickname = RE.Regex.Replace(customerName, @"[^\w-]", "", RE.RegexOptions.None, TimeSpan.FromSeconds(1.5)) + "-" + customerNo + "-Kund";
            if (customerType == "Supplier")
                mailNickname = RE.Regex.Replace(customerName, @"[^\w-]", "", RE.RegexOptions.None, TimeSpan.FromSeconds(1.5)) + "-" + customerNo + "-Lev";

            illegalChars.ForEach(c => mailNickname = mailNickname.Replace(c.ToString(), ""));

            return mailNickname.Replace("é", "e");
        }

        public FindCustomerGroupResult FindCustomerGroupAndDrive(string customerName, string customerNo, string customerType)
        {
            FindCustomerGroupResult returnValue = new FindCustomerGroupResult();
            returnValue.Success = false;
            string mailNickname = this.GetMailNickname(customerName, customerNo, customerType);

            telemetryClient.TrackEvent(new EventTelemetry($"Trying to get group for {mailNickname}."));
            try
            {
                FindCustomerResult findCustomer = GetCustomer(customerNo, customerType, customerName);

                if (findCustomer.Success)
                {
                    returnValue.customer = findCustomer.customer;

                    if (returnValue.customer != null)
                    {
                        returnValue = FindCustomerGroupAndDrive(returnValue.customer).Result;
                    }
                }
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to get group and drive for {mailNickname} with error: " + ex.ToString()));
            }

            return returnValue;
        }

        public async Task<FindCustomerGroupResult> FindCustomerGroupAndDrive(Customer customer)
        {
            FindCustomerGroupResult returnValue = new FindCustomerGroupResult();
            returnValue.Success = false;
            Drive groupDrive = null;
            List<DriveItem> rootItems = new List<DriveItem>();
            List<DriveItem> generalItems = new List<DriveItem>();
            FindGroupResult result = new FindGroupResult() { Success = false };

            telemetryClient.TrackEvent(new EventTelemetry($"Trying to get group drive for {customer.Name}."));
            if (!string.IsNullOrEmpty(customer.GroupID))
            {
                result = await msGraph.GetGroupById(customer.GroupID);
            }
            else
            {
                telemetryClient.TrackEvent(new EventTelemetry($"Missing group id in database so trying mailnickname {customer.Name}."));
                string mailNickname = this.GetMailNickname(customer.Name, customer.ExternalId, customer.Type);
                result = await msGraph.FindGroupByName(mailNickname, false);
            }

            if (result.Success)
            {
                returnValue.customer = customer;

                if (result.Count > 1)
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"Found multiple groups for {customer.Name}. Returning first match."));
                    returnValue.group = result.groups[0];
                    groupDrive = await msGraph.GetGroupDrive(result.groups[0]);
                }
                else
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"Found group for {customer.Name}."));
                    returnValue.group = result.group;
                    groupDrive = await msGraph.GetGroupDrive(result.group);
                }

                returnValue.customer.GroupID = returnValue.group.Id;
                returnValue.customer.GroupCreated = true;
                UpdateCustomer(returnValue.customer, "group info");

                if (groupDrive != default(Drive))
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"Found group drive for {customer.Name}."));
                    returnValue.Success = true;
                    returnValue.groupDrive = groupDrive;
                    returnValue.customer.DriveID = returnValue.groupDrive.Id;
                    UpdateCustomer(returnValue.customer, "drive info");
                    rootItems = await msGraph.GetDriveRootItems(groupDrive);

                    if (rootItems.Count > 0)
                    {
                        telemetryClient.TrackEvent(new EventTelemetry($"Fetched root items in group drive for {customer.Name}."));
                        returnValue.rootItems = rootItems;
                    }
                }

                if (rootItems.Count > 0)
                {
                    var generalFolder = rootItems.FirstOrDefault(ri => ri.Name == "General");

                    if (generalFolder != default(DriveItem))
                    {
                        telemetryClient.TrackEvent(new EventTelemetry($"Fetched general folder in group drive for {customer.Name}."));
                        returnValue.generalFolder = generalFolder;
                        returnValue.customer.GeneralFolderID = generalFolder.Id;
                        returnValue.customer.GeneralFolderCreated = true;
                        UpdateCustomer(returnValue.customer, "general folder info");
                    }
                }
            }

            return returnValue;
        }

        public FindOrderGroupAndFolder GetOrderGroupAndFolder(string OrderNo)
        {
            FindOrderGroupAndFolder returnValue = new FindOrderGroupAndFolder();
            returnValue.Success = false;
            Order order = this.GetOrderFromCDN(OrderNo);
            telemetryClient.TrackEvent(new EventTelemetry($"Trying to fetch CDN item for {OrderNo}."));
            returnValue = this.GetOrderGroupAndFolder(order).Result;

            return returnValue;
        }

        public async Task<FindOrderGroupAndFolder> GetOrderGroupAndFolder(Order order)
        {
            FindOrderGroupAndFolder returnValue = new FindOrderGroupAndFolder();
            returnValue.Success = false;

            if (order != null && !string.IsNullOrEmpty(order.CustomerNo) && !string.IsNullOrEmpty(order.CustomerType))
            {
                FindCustomerResult customerName = GetCustomer(order);

                if (customerName.Success && customerName.customer != null)
                {
                    telemetryClient.TrackEvent(new EventTelemetry($"Got customer name from cdn for {customerName.customer.Name}."));
                    order.Customer = customerName.customer;
                    returnValue.customer = customerName.customer;

                    FindCustomerGroupResult findCustomerGroupResult = new FindCustomerGroupResult();

                    if (returnValue.customer != null && string.IsNullOrEmpty(returnValue.customer.GroupID) || string.IsNullOrEmpty(returnValue.customer.DriveID))
                    {
                        findCustomerGroupResult = this.FindCustomerGroupAndDrive(returnValue.customer.Name, returnValue.customer.ExternalId, returnValue.customer.Type);
                    }
                    else
                    {
                        findCustomerGroupResult = await this.FindCustomerGroupAndDrive(returnValue.customer);
                    }

                    if (findCustomerGroupResult.Success)
                    {
                        returnValue.customer.GroupID = findCustomerGroupResult.group.Id;
                        returnValue.customer.DriveID = findCustomerGroupResult.groupDrive.Id;
                        returnValue.customer.GroupCreated = true;
                        UpdateCustomer(returnValue.customer, "group and drive info");
                        telemetryClient.TrackEvent(new EventTelemetry($"Found group for {returnValue.customer.Name} and order {order.ExternalId}."));

                        try
                        {
                            returnValue.orderTeam = await graphClient.Groups[findCustomerGroupResult.group.Id].Team.GetAsync();

                            if (returnValue.orderTeam != null)
                            {
                                returnValue.customer.TeamCreated = true;
                                returnValue.customer.TeamID = returnValue.orderTeam.Id;
                            }
                            UpdateCustomer(returnValue.customer, "team info");
                            telemetryClient.TrackEvent(new EventTelemetry($"Found team for {returnValue.customer.Name} and order {order.ExternalId}."));
                        }
                        catch (Exception ex)
                        {
                            telemetryClient.TrackException(ex);
                            telemetryClient.TrackTrace(new TraceTelemetry($"Failed to find team for {returnValue.customer.Name} and order {order.ExternalId}."));
                        }

                        returnValue.Success = true;
                        returnValue.orderGroup = findCustomerGroupResult.group;
                        returnValue.orderDrive = findCustomerGroupResult.groupDrive;

                        if (findCustomerGroupResult.generalFolder != null)
                        {
                            telemetryClient.TrackEvent(new EventTelemetry($"Found general folder for {returnValue.customer.Name} and order {order.ExternalId}."));
                            returnValue.generalFolder = findCustomerGroupResult.generalFolder;
                            returnValue.customer.GeneralFolderCreated = true;
                            returnValue.customer.GeneralFolderID = returnValue.generalFolder.Id;
                            UpdateCustomer(returnValue.customer, "general folder info");
                        }

                        if (returnValue.customer.GeneralFolderCreated)
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
                                        telemetryClient.TrackEvent(new EventTelemetry($"Changed order no for quote: {order.ExternalId} to: {orderMatch.Value}"));
                                        order.ExternalId = orderMatch.Value;
                                    }

                                    break;
                                case "Offer":
                                    parentName = "Offert";
                                    RE.Match offerMatch = RE.Regex.Match(order.ExternalId, @"^([A-Z]?\d+)");

                                    if (offerMatch.Success)
                                    {
                                        telemetryClient.TrackEvent(new EventTelemetry($"Changed order no for quote: {order.ExternalId} to: {offerMatch.Value}"));
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
                                DriveItem foundOrderFolder = await msGraph.FindItem(returnValue.orderDrive, "General/" + parentName + "/" + order.ExternalId, false);

                                if (foundOrderFolder != null)
                                {
                                    telemetryClient.TrackEvent(new EventTelemetry($"Found order folder for {order.ExternalId} in customer/supplier {returnValue.customer.Name}."));
                                    returnValue.orderFolder = foundOrderFolder;
                                    order.CreatedFolder = true;
                                    order.CustomerID = returnValue.customer.ID;
                                    order.GroupFound = true;
                                    order.GeneralFolderFound = true;
                                    order.FolderID = returnValue.orderFolder.Id;
                                    order.OrdersFolderFound = true;
                                    UpdateOrder(order, "folder info");
                                }
                                else
                                {
                                    List<DriveItem> rootItems = await msGraph.GetDriveRootItems(returnValue.orderDrive);

                                    foreach(DriveItem rootItem in rootItems)
                                    {
                                        if(rootItem.Name == "General")
                                        {
                                            List<DriveItem> generalItems = await msGraph.GetDriveFolderChildren(returnValue.orderDrive, rootItem, false);

                                            foreach (DriveItem generalItem in generalItems)
                                            {
                                                if (generalItem.Name == parentName)
                                                {
                                                    List<DriveItem> folderItems = await msGraph.GetDriveFolderChildren(returnValue.orderDrive, generalItem, false);

                                                    foreach(DriveItem folderItem in folderItems)
                                                    {
                                                        if(folderItem.Name == order.ExternalId)
                                                        {
                                                            telemetryClient.TrackEvent(new EventTelemetry($"Found order folder for {order.ExternalId} in customer/supplier {returnValue.customer.Name}."));
                                                            returnValue.orderFolder = folderItem;
                                                            order.CreatedFolder = true;
                                                            order.CustomerID = returnValue.customer.ID;
                                                            order.GroupFound = true;
                                                            order.GeneralFolderFound = true;
                                                            order.FolderID = returnValue.orderFolder.Id;
                                                            order.OrdersFolderFound = true;
                                                            UpdateOrder(order, "folder info");
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
                                telemetryClient.TrackException(ex);
                                telemetryClient.TrackTrace($"Failed to get folder for order {order.ExternalId}.");
                            }
                        }
                    }
                }
            }

            return returnValue;
        }

        public string FindOrderNoInString(string input)
        {
            string returnValue = "";
            RE.Match orderMatches = RE.Regex.Match(input, @"B\d{6}|T\d{5}|A\d{5}|Z\d{5}|G\d{4}|R\d{2}|E\d{5,7}|F\d{5}|H\d{5,6}|K\d{5,6}|\d{5}-\d{2}|Q\d{5}-\d{2}");

            if (orderMatches.Success)
            {
                returnValue = orderMatches.Value;
            }

            return returnValue;
        }

        public string FindCustomerNoInString(string input)
        {
            string returnValue = "";
            RE.Match orderMatches = RE.Regex.Match(input, @"(\d+)(\s?|-?)");

            if (orderMatches.Success)
            {
                returnValue = orderMatches.Groups[1].Value;
            }

            return returnValue;
        }

        public List<Order> GetUnhandledOrderItems()
        {
            List<Order> returnValue = new List<Order>();
            returnValue = Services.ExecSQLQuery<Order>("SELECT * FROM Orders WHERE Handled = 0", new Dictionary<string, object>(), SqlConnectionString);
            return returnValue;
        }

        public async Task<DriveItem> GetEmailsFolder(string parent, string month, string year)
        {
            Drive groupDrive = default(Drive);

            try
            {
                groupDrive = await msGraph.GetGroupDrive(CDNTeamID);
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to get drive for CDN Team with error: " + ex.ToString()));
            }

            DriveItem emailFolder = default(DriveItem);

            if (groupDrive != default(Drive))
            {
                try
                {
                    emailFolder = await msGraph.FindItem(groupDrive, parent + "/EmailMessages_" + month + "_" + year, false);
                }
                catch (Exception ex)
                {
                    telemetryClient.TrackException(ex);
                    telemetryClient.TrackTrace(new TraceTelemetry($"Failed to get email folder for CDN Team with error: " + ex.ToString()));
                }
            }

            return emailFolder;
        }

        public async Task<DriveItem> GetGeneralFolder(string groupId)
        {
            Drive groupDrive = default(Drive);

            try
            {
                groupDrive = await msGraph.GetGroupDrive(groupId);
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to get drive for group {groupId} with error: " + ex.ToString()));
            }

            DriveItem generalFolder = default(DriveItem);

            if (groupDrive != default(Drive))
            {
                try
                {
                    generalFolder = await msGraph.FindItem(groupDrive, "General", false);
                }
                catch (Exception ex)
                {
                    telemetryClient.TrackException(ex);
                    telemetryClient.TrackTrace(new TraceTelemetry($"Failed to get general folder in group {groupId} with error: " + ex.ToString()));
                }
            }

            return generalFolder;
        }

        public async Task<CreateCustomerResult> CreateCustomerGroup(Customer customer)
        {
            CreateCustomerResult returnValue = new CreateCustomerResult();
            returnValue.Success = false;
            Group group = default(Group);
            string[] admins = config["Admins"].Split(',');
            List<string> adminids = new List<string>();
            string mailNickname = this.GetMailNickname(customer.Name, customer.ExternalId, customer.Type);
            adminids = await GetAdmins(new Customer(), admins);
            string GroupName = "";

            if (customer.Type == "Customer")
                GroupName = customer.Name + " (" + customer.ExternalId + ") - Kund";
            if (customer.Type == "Supplier")
                GroupName = customer.Name + " (" + customer.ExternalId + ") - Lev";

            try
            {
                //Create a group without owners
                group = await msGraph.CreateGroup(GroupName, mailNickname, adminids);
                telemetryClient.TrackEvent(new EventTelemetry($"Created group for customer {customer.Name} ({customer.ExternalId})"));
            }
            catch (Exception ex)
            {
                telemetryClient.TrackException(ex);
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create group for {customer.Name} ({customer.ExternalId}) with error: " + ex.ToString()));
            }

            //if the group was created
            if(group != null)
            {
                customer.GroupID = group.Id;

                //get the group drive (will probably fail since thr group takes a while to create)
                try
                {
                    Drive groupDrive = await msGraph.GetGroupDrive(group);

                    if (groupDrive != null)
                    {
                        customer.DriveID = groupDrive.Id;
                    }
                }
                catch (Exception ex)
                {
                    log.LogInformation(ex.ToString());
                }

                returnValue.group = group;
                returnValue.customer = customer;
                returnValue.Success = true;
            }
            else
            {
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create group {customer.Name} ({customer.ExternalId})"));
            }

            return returnValue;
        }

        /// <summary>
        /// Create group and team for customer or supplier
        /// No checking is done if group or team already exists.
        /// </summary>
        /// <param name="customer"></param>
        /// <returns></returns>
        public async Task<CreateCustomerResult> CreateCustomerOrSupplier(Customer customer)
        {
            CreateCustomerResult returnValue = new CreateCustomerResult();
            returnValue.Success = false;
            Group group = default(Group);
            string[] admins = config["Admins"].Split(',');
            List<string> adminids = new List<string>();
            string mailNickname = this.GetMailNickname(customer.Name, customer.ExternalId, customer.Type);
            adminids = await GetAdmins(customer, admins);
            string GroupName = "";

            if (customer.Type == "Customer")
                GroupName = customer.Name + " (" + customer.ExternalId + ") - Kund";
            if (customer.Type == "Supplier")
                GroupName = customer.Name + " (" + customer.ExternalId + ") - Lev";

            //find group if it exists or try to create it
            if(customer?.GroupID != null && customer?.GroupID != string.Empty)
            {
                var findGroup = await msGraph.GetGroupById(customer.GroupID);

                if (findGroup.Success)
                {
                    group = findGroup.group;
                }
            }
            else
            {
                try
                {
                    group = await msGraph.CreateGroup(GroupName, mailNickname, adminids);
                    telemetryClient.TrackEvent(new EventTelemetry($"Created group for customer {customer.Name} ({customer.ExternalId})"));
                }
                catch (Exception ex)
                {
                    telemetryClient.TrackException(ex);
                    telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create group for {customer.Name} ({customer.ExternalId}) with error: " + ex.ToString()));
                }
            }

            if (group != null)
            {
                customer.GroupID = group.Id;

                try
                {
                    Drive groupDrive = await msGraph.GetGroupDrive(group);

                    if (groupDrive != null)
                    {
                        customer.DriveID = groupDrive.Id;
                    }
                }
                catch (Exception ex)
                {
                    telemetryClient.TrackException(ex);
                }

                customer.GroupCreated = true;
                UpdateCustomer(customer, "group and drive info");

                var team = await msGraph.CreateTeamFromGroup(group);
                telemetryClient.TrackEvent(new EventTelemetry($"Created team for {customer.Name} ({customer.ExternalId})"));

                if (team != null)
                {
                    customer.TeamCreated = true;
                    customer.TeamID = team.Id;
                    customer.TeamUrl = team.WebUrl;
                    UpdateCustomer(customer, "team info");

                    try
                    {
                        string ContentUrl = "https://holtabcustomercard.azurewebsites.net/Home/Index?id=" + team.Id;
                        var groupDrive = await msGraph.GetGroupDrive(group.Id);

                        if (groupDrive != null)
                        {
                            var root = await graphClient.Drives[groupDrive.Id].Root.GetAsync();

                            if(root != null)
                            {
                                var channels = await graphClient.Teams[team.Id].Channels.GetAsync();

                                if(channels?.Value?.Count > 0)
                                {
                                    var app = await msGraph.AddTeamApp(team, "e2cb3981-47e7-47b3-a0e1-f9078d342253");
                                    await msGraph.AddChannelApp(team, app, channels.Value[0], "Om Företaget", System.Guid.NewGuid().ToString("D").ToUpperInvariant(), ContentUrl, root.WebUrl, null);
                                    telemetryClient.TrackEvent(new EventTelemetry($"Installed teams app for {customer.Name} ({customer.ExternalId})"));
                                    customer.InstalledApp = true;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        telemetryClient.TrackException(ex);
                        telemetryClient.TrackTrace(new TraceTelemetry($"Failed to install teams app for {customer.Name} with error: " + ex.ToString()));
                    }

                    UpdateCustomer(customer, "team app info");

                    returnValue.group = group;
                    returnValue.team = team;
                    returnValue.customer = customer;
                    returnValue.Success = true;
                }
                else
                {
                    telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create team for group {customer.Name} ({customer.ExternalId})"));
                }
            }
            else
            {
                telemetryClient.TrackTrace(new TraceTelemetry($"Failed to create group {customer.Name} ({customer.ExternalId})"));
            }

            return returnValue;
        }

        /// <summary>
        /// Copy files and folders from structure in CDN site to group site for customer or supplier
        /// Expects the general folder to exist before the function is run
        /// </summary>
        /// <param name="customer"></param>
        /// <returns></returns>
        public async Task<bool> CopyRootStructure(Customer customer)
        {
            bool returnValue = false;
            var cdnDrive = await msGraph.GetSiteDrive(cdnSiteId);

            if (cdnDrive != null)
            {
                DriveItem source = default(DriveItem);

                try
                {
                    if (customer.Type == "Customer")
                    {
                        source = await msGraph.FindItem(cdnDrive, "Dokumentstruktur Kund", false);
                    }
                    else if (customer.Type == "Supplier")
                    {
                        source = await msGraph.FindItem(cdnDrive, "Dokumentstruktur Leverantör", false);
                    }
                }
                catch (Exception ex)
                {
                    telemetryClient.TrackException(ex);
                    telemetryClient.TrackTrace(new TraceTelemetry($"Failed to get templates for {customer.Name} with error: " + ex.ToString()));
                }

                telemetryClient.TrackEvent(new EventTelemetry($"Found CDN folder structure template for {customer.Name} ({customer.ExternalId})"));

                if (source != default(DriveItem))
                {
                    DriveItem generalFolder = default(DriveItem);

                    if (string.IsNullOrEmpty(customer.GeneralFolderID))
                    {
                        try
                        {
                            generalFolder = await this.GetGeneralFolder(customer.GroupID);

                            if (generalFolder != null)
                            {
                                customer.GeneralFolderID = generalFolder.Id;
                                telemetryClient.TrackEvent(new EventTelemetry($"Found general folder for {customer.Name} ({customer.ExternalId})"));
                            }
                        }
                        catch (Exception ex)
                        {
                            telemetryClient.TrackException(ex);
                            telemetryClient.TrackTrace(new TraceTelemetry($"Failed to get general folder for {customer.Name} with error: " + ex.ToString()));
                        }
                    }

                    if (!string.IsNullOrEmpty(customer.GeneralFolderID))
                    {
                        try
                        {
                            var children = await msGraph.GetDriveFolderChildren(cdnDrive, source, true);

                            foreach (var child in children)
                            {
                                await msGraph.CopyFolder(customer.GroupID, customer.GeneralFolderID, child, true, false);
                            }

                            telemetryClient.TrackEvent(new EventTelemetry($"Copied templates for {customer.Name} ({customer.ExternalId})"));
                            returnValue = true;
                        }
                        catch (Exception ex)
                        {
                            telemetryClient.TrackException(ex);
                            telemetryClient.TrackTrace($"Failed to copy template structure for {customer.Name} ({customer.ExternalId})");
                        }

                    }
                }
            }

            return returnValue;
        }

        public async Task<List<string>> GetAdmins(Customer customer, string[] admins)
        {
            List<string> adminids = new List<string>();
            List<string> _admins = new List<string>();
            _admins.AddRange(admins);

            //if seller exists add it to admins list
            if (!String.IsNullOrEmpty(customer.Seller) && !_admins.Exists(a => a == customer.Seller))
                _admins.Add(customer.Seller);

            //Get all admin ids
            foreach (string user in admins)
            {
                try
                {
                    var graphUser = await graphClient.Users[user].GetAsync();

                    if (graphUser != null)
                    {
                        adminids.Add("https://graph.microsoft.com/v1.0/users/" + graphUser.Id);
                    }
                    else
                    {
                        telemetryClient.TrackTrace($"Failed to find user {user}");
                    }
                }
                catch (Exception ex)
                {
                    telemetryClient.TrackException(ex);
                    telemetryClient.TrackTrace($"Failed to get user {user}" + ex.ToString());
                }
            }

            return adminids;
        }


    }
}
