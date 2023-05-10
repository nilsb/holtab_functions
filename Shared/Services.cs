using Microsoft.Data.SqlClient;
using Shared.Models;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;

namespace Shared
{
    public class Services
    {
        public readonly bool init;
        private readonly string SqlConnectionString;

        public Services(string? _sqlConnectionString)
        {
            this.init = false;

            if (_sqlConnectionString != null)
            {
                this.SqlConnectionString = _sqlConnectionString;
                this.init = true;
            }
            else
            {
                this.SqlConnectionString = "";
            }
        }

        //SELECT commands
        private readonly string SelectOrderByIDCommand = "SELECT * FROM Orders WHERE ID = @ID";
        private readonly string SelectOrderByExternalIDCommand = "SELECT * FROM Orders WHERE ExternalId = @ExternalId";
        private readonly string SelectCustomerByIDCommand = "SELECT * FROM Customers WHERE ID = @CustomerID";
        private readonly string SelectCustomerByExternalIDCommand = "SELECT * FROM Customers WHERE ExternalId = @ExternalId AND [Type] = @Type";

        #region SelectOrder
        public Order? GetOrderFromDB(string orderNo)
        {
            Order? returnValue = null;
            Dictionary<string, object> keys = new Dictionary<string, object>();
            keys.Add("ExternalId", orderNo);
            List<Order> result = ExecSQLQuery<Order>(SelectOrderByExternalIDCommand, keys);

            if (result.Count > 0)
            {
                result.ForEach(or =>
                {
                    if (or.Customer != null)
                    {
                        or.CustomerID = or.Customer.ID;
                    }
                    else
                    {
                        Customer? dbCustomer = GetCustomerFromDB(or.CustomerID);

                        if (dbCustomer != null)
                        {
                            or.Customer = dbCustomer;
                        }
                    }
                });

                returnValue = result[0];
            }

            return returnValue;
        }

        public Order? GetOrderFromDB(Guid ID)
        {
            Order? returnValue = null;
            Dictionary<string, object> keys = new Dictionary<string, object>();
            keys.Add("OrderID", ID);
            List<Order> result = ExecSQLQuery<Order>(SelectOrderByIDCommand, keys);

            if (result.Count > 0)
            {
                result.ForEach(or =>
                {
                    or.Customer = GetCustomerFromDB(or.CustomerID);

                    if (or.Customer != null)
                    {
                        or.CustomerID = or.Customer.ID;
                    }
                });

                returnValue = result[0];
            }

            return returnValue;
        }
        #endregion

        #region AddOrder
        public bool AddOrderInDB(Order obj)
        {
            return InsertSQLQuery(obj, "Orders");
        }
        #endregion

        #region OrderUpdate
        public bool UpdateOrderInDB(Order obj)
        {
            Dictionary<string, object> keys = new Dictionary<string, object>();

            if (obj.ID != Guid.Empty)
            {
                keys.Add("ID", obj.ID);
            }
            else
            {
                keys.Add("ExternalId", obj.ExternalId);
                keys.Add("Type", obj.Type);
            }

            return UpdateSQLQuery(obj, "Orders", keys);
        }
        #endregion

        #region SelectCustomer
        public List<Customer> GetCustomerFromDB(string? customerNo, string? customerType)
        {
            List<Customer> returnValue = new List<Customer>();

            if(!string.IsNullOrEmpty(customerNo) && !string.IsNullOrEmpty(customerType))
            {
                Dictionary<string, object> keys = new Dictionary<string, object>
                {
                    { "ExternalId", customerNo },
                    { "Type", customerType }
                };

                List<Customer> ret = ExecSQLQuery<Customer>(SelectCustomerByExternalIDCommand, keys);

                if (ret.Count > 0)
                {
                    returnValue = ret;
                }
            }

            return returnValue;

        }

        public Customer? GetCustomerFromDB(Guid ID)
        {
            Customer? returnValue = null;
            Dictionary<string, object> keys = new Dictionary<string, object>();
            keys.Add("CustomerID", ID);

            List<Customer> ret = ExecSQLQuery<Customer>(SelectCustomerByIDCommand, keys);

            if (ret.Count > 0)
            {
                returnValue = ret[0];
            }

            return returnValue;

        }
        #endregion

        #region CustomerUpdate
        public bool UpdateCustomerInDB(Customer obj)
        {
            Dictionary<string, object> keys = new Dictionary<string, object>();

            if (obj.ID != Guid.Empty)
            {
                keys.Add("ID", obj.ID);
            }
            else 
            {
                keys.Add("ExternalId", obj.ExternalId);
            }


            return UpdateSQLQuery(obj, "Customers", keys);
        }
        #endregion

        #region CustomerAdd
        public bool AddCustomerInDB(Customer obj)
        {
            return InsertSQLQuery(obj, "Customers");
        }
        #endregion

        #region Supporting
        public bool Log(string message)
        {
            int affectedRows = 0;

            using (SqlConnection conn = new SqlConnection(this.SqlConnectionString))
            {
                using (SqlCommand command = new SqlCommand())
                {
                    conn.Open();
                    command.CommandText = "INSERT INTO Debug (message, date) VALUES (@message, GETDATE())";
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    command.Parameters.AddWithValue("message", message);
                    affectedRows = command.ExecuteNonQuery();
                    conn.Close();
                }
            }

            return (affectedRows > 0);

        }

        public string GetUpdateSQLQuery<T>(T? src, string tablename, Dictionary<string, object>? keys)
        {
            if (src == null || string.IsNullOrEmpty(tablename))
                return string.Empty;

            string query = "UPDATE " + tablename + " SET ";

            //Set all mapped properties of the update query
            var properties = src.GetType().GetProperties();
            List<PropertyInfo> mappedProperties = new List<PropertyInfo>();
            mappedProperties.AddRange(properties);

            foreach (PropertyInfo prop in mappedProperties.Where(prop => !Attribute.IsDefined(prop, typeof(NotMappedAttribute)) && !Attribute.IsDefined(prop, typeof(KeyAttribute))))
            {
                if(prop.GetValue(src) != null)
                {
                    query += prop.Name + "=@" + prop.Name + ", ";
                }
            }

            query = query.TrimEnd(' ');
            query = query.TrimEnd(',');

            if(keys != null && keys?.Count > 0)
            {
                query += " WHERE ";

                //Add conditions for update
                foreach (var kvp in keys)
                {
                    if (kvp.Value != null)
                    {
                        if (kvp.Value is int)
                        {
                            query += kvp.Key + " = " + kvp.Value.ToString();
                        }
                        else if (kvp.Value is bool)
                        {
                            if ((bool)kvp.Value == true)
                            {
                                query += kvp.Key + " = 1";
                            }
                            else
                            {
                                query += kvp.Key + " = 0";
                            }
                        }
                        else if (kvp.Value is DateTime)
                        {
                            query += "'" + ((DateTime)kvp.Value).ToString("yyyy-MM-dd HH:mm:ss") + "'";
                        }
                        else
                        {
                            query += kvp.Key + " = '" + kvp.Value.ToString() + "'";
                        }
                    }
                    else
                    {
                        query += kvp.Key + " = NULL";
                    }

                    query += " AND ";
                }
                query = query.TrimEnd(' ');
                query = query.TrimEnd('D');
                query = query.TrimEnd('N');
                query = query.TrimEnd('A');
                query = query.TrimEnd(' ');
            }

            return query;
        }

        /// <summary>
        /// Update an object in the database.
        /// This will overwrite all the properties in the database so make sure you fetch the complete object before updating.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src"></param>
        /// <param name="tablename"></param>
        /// <param name="keys"></param>
        /// <param name="_config"></param>
        /// <returns></returns>
        public bool UpdateSQLQuery<T>(T? src, string tablename, Dictionary<string, object> keys)
        {
            bool returnValue = false;
            int affectedRows = 0;
            string query = GetUpdateSQLQuery(src, tablename, keys);

            //Set all mapped properties of the update query
            var properties = src?.GetType().GetProperties();
            List<PropertyInfo> mappedProperties = new List<PropertyInfo>(); 
            
            if(properties != null)
            {
                mappedProperties.AddRange(properties);
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(this.SqlConnectionString))
                {
                    using (SqlCommand command = new SqlCommand())
                    {
                        conn.Open();
                        command.CommandText = query;
                        command.CommandType = System.Data.CommandType.Text;
                        command.Connection = conn;

                        //add parameters for conditions
                        if (keys.Count > 0)
                        {
                            foreach (var kvp in keys)
                            {
                                command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }

                        //add parameters for values to set
                        foreach (PropertyInfo prop in mappedProperties.Where(prop => !Attribute.IsDefined(prop, typeof(NotMappedAttribute)) && !Attribute.IsDefined(prop, typeof(KeyAttribute))))
                        {
                            var value = prop.GetValue(src);

                            if (!keys.ContainsKey(prop.Name) && value != null)
                            {
                                command.Parameters.AddWithValue(prop.Name, value);
                            }
                        }

                        affectedRows = command.ExecuteNonQuery();
                        conn.Close();

                        returnValue = (affectedRows > 0);
                    }
                }
            }
            catch (Exception)
            {
            }

            return returnValue;
        }

        /// <summary>
        /// Get the sql insert query as a string for an object. (Intended for internal use)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src"></param>
        /// <param name="tablename"></param>
        /// <param name="keys"></param>
        /// <returns></returns>
        public string GetSQLInsertQuery<T>(T src, string tablename)
        {
            string query = "INSERT INTO " + tablename + " (";

            //Set all mapped properties of the update query
            var properties = src.GetType().GetProperties();
            List<PropertyInfo> mappedProperties = new List<PropertyInfo>();
            mappedProperties.AddRange(properties);

            foreach (PropertyInfo prop in mappedProperties.Where(prop => !Attribute.IsDefined(prop, typeof(NotMappedAttribute))))
            {
                query += "[" + prop.Name + "], ";
            }

            query = query.TrimEnd(' ');
            query = query.TrimEnd(',');
            query += ") VALUES (";

            //Add values for insert
            foreach (PropertyInfo prop in mappedProperties.Where(prop => !Attribute.IsDefined(prop, typeof(NotMappedAttribute))))
            {
                var value = prop.GetValue(src);

                if (value != null)
                {
                    if (value is int || value is long || value is double)
                    {
                        if(value is int)
                        {
                            query += ((int)value).ToString().Replace(",", ".");
                        }
                        if(value is long)
                        {
                            query += ((long)value).ToString().Replace(",", ".");
                        }
                        if(value is double)
                        {
                            query += ((double)value).ToString().Replace(",", ".");
                        }
                    }
                    else if (value is bool)
                    {
                        if ((bool)value == true)
                        {
                            query += "1";
                        }
                        else
                        {
                            query += "0";
                        }
                    }
                    else if (value is DateTime time)
                    {
                        query += "'" + time.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                    }
                    else
                    {
                        query += "'" + value.ToString() + "'";
                    }
                }
                else
                {
                    query += "NULL";
                }

                query += ", ";
            }

            query = query.TrimEnd(' ');
            query = query.TrimEnd(',');
            query += ")";

            return query;
        }

        /// <summary>
        /// Insert an object into the database.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src"></param>
        /// <param name="tablename"></param>
        /// <param name="_config"></param>
        /// <returns></returns>
        public bool InsertSQLQuery<T>(T src, string tablename)
        {
            bool returnValue = false;
            int affectedRows = 0;
            string query = GetSQLInsertQuery(src, tablename);

            try
            {
                using (SqlConnection conn = new SqlConnection(this.SqlConnectionString))
                {
                    using (SqlCommand command = new SqlCommand())
                    {
                        conn.Open();
                        command.CommandText = query;
                        command.CommandType = System.Data.CommandType.Text;
                        command.Connection = conn;
                        affectedRows = command.ExecuteNonQuery();
                        conn.Close();

                        returnValue = (affectedRows > 0);
                    }
                }

            }
            catch (Exception)
            {
            }

            return returnValue;
        }

        public bool InsertSQLQuery(string query)
        {
            bool returnValue = false;
            int affectedRows = 0;

            try
            {
                using (SqlConnection conn = new SqlConnection(this.SqlConnectionString))
                {
                    using (SqlCommand command = new SqlCommand())
                    {
                        conn.Open();
                        command.CommandText = query;
                        command.CommandType = System.Data.CommandType.Text;
                        command.Connection = conn;
                        affectedRows = command.ExecuteNonQuery();
                        conn.Close();

                        returnValue = (affectedRows > 0);
                    }
                }
            }
            catch (Exception)
            {
            }

            return returnValue;
        }

        /// <summary>
        /// Fetch a list of matching objects from the database.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <param name="keys"></param>
        /// <param name="_config"></param>
        /// <returns></returns>
        public List<T> ExecSQLQuery<T>(string query, Dictionary<string, object> keys)
        {
            List<T> list = new List<T>();

            try
            {
                using (SqlConnection conn = new SqlConnection(this.SqlConnectionString))
                {
                    using (SqlCommand command = new SqlCommand())
                    {
                        conn.Open();
                        command.CommandText = query;
                        command.CommandType = System.Data.CommandType.Text;
                        command.Connection = conn;

                        foreach (var kvp in keys)
                        {
                            command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                        }

                        using (var result = command.ExecuteReader())
                        {
                            T? obj = default(T);

                            while (result.Read())
                            {
                                obj = Activator.CreateInstance<T>();

                                if (obj != null)
                                {
                                    foreach (PropertyInfo prop in obj.GetType().GetProperties())
                                    {
                                        if (result.GetColumnSchema().Any(c => c.ColumnName == prop.Name))
                                        {
                                            if (!object.Equals(result[prop.Name], DBNull.Value))
                                            {
                                                try
                                                {
                                                    prop.SetValue(obj, result[prop.Name], null);
                                                }
                                                catch (Exception)
                                                {
                                                }
                                            }
                                        }
                                    }

                                    list.Add(obj);
                                }
                            }
                        }

                        conn.Close();
                    }
                }
            }
            catch (Exception)
            {
            }

            return list;
        }

        /// <summary>
        /// Execute a query not expected to return any rows.
        /// </summary>
        /// <param name="query"></param>
        /// <param name="keys"></param>
        /// <param name="_config"></param>
        /// <returns></returns>
        public bool ExecSQLNonQuery(string query, Dictionary<string, object> keys)
        {
            bool returnValue = false;
            int affectedRows = 0;

            try
            {
                using (SqlConnection conn = new SqlConnection(this.SqlConnectionString))
                {
                    using (SqlCommand command = new SqlCommand())
                    {
                        conn.Open();
                        command.CommandText = query;
                        command.CommandType = System.Data.CommandType.Text;
                        command.Connection = conn;

                        foreach (var kvp in keys)
                        {
                            command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                        }

                        affectedRows = command.ExecuteNonQuery();
                        conn.Close();

                        returnValue = (affectedRows > 0);
                    }
                }
            }
            catch (Exception)
            {
            }

            return returnValue;
        }
        #endregion
    }
}
