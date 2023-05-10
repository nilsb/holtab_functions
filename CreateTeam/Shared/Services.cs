using CreateTeam.Models;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Reflection;

namespace CreateTeam.Shared
{
    public static class Services
    {
        private static readonly string SqlConnectionString = "Server=tcp:holtab.database.windows.net,1433;Initial Catalog=holtab_integration_db;Persist Security Info=False;User ID=holtabdbadmin;Password=DVKvXWmS#LJp&EvE!!6yXRDWL&JX$3##wMNio9DXd9jxkA^h9w$pBBKvQidu3rSjyA83%SNk9EWTjZAojW%YE^WJHzgwJQfJ*ALH!8dt%zyvWyGE8MMucthByU$uVJk5;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";

        //SELECT commands
        private static readonly string SelectOrderByIDCommand = "SELECT * FROM Orders WHERE ID = @ID";
        private static readonly string SelectOrderByExternalIDCommand = "SELECT * FROM Orders WHERE ExternalId = @ExternalId";
        private static readonly string SelectCustomerByIDCommand = "SELECT * FROM Customers WHERE ID = @CustomerID";
        private static readonly string SelectCustomerByExternalIDCommand = "SELECT * FROM Customers WHERE ExternalId = @ExternalId AND [Type] = @Type";

        #region SelectOrder
        public static Order GetOrderFromDB(string orderNo, string connstr)
        {
            Order returnValue = null;
            Dictionary<string, object> keys = new Dictionary<string, object>();
            keys.Add("ExternalId", orderNo);
            List<Order> result = ExecSQLQuery<Order>(SelectOrderByExternalIDCommand, keys, connstr);

            if (result.Count > 0)
            {
                result.ForEach(or =>
                {
                    or.Customer = GetCustomerFromDB(or.CustomerID, connstr);

                    if(or.Customer != null)
                    {
                        or.CustomerID = or.Customer.ID;
                    }
                });

                returnValue = result[0];
            }

            return returnValue;
        }

        public static Order GetOrderFromDB(Guid ID, string connstr)
        {
            Order returnValue = null;
            Dictionary<string, object> keys = new Dictionary<string, object>();
            keys.Add("OrderID", ID);
            List<Order> result = ExecSQLQuery<Order>(SelectOrderByIDCommand, keys, connstr);

            if (result.Count > 0)
            {
                result.ForEach(or =>
                {
                    or.Customer = GetCustomerFromDB(or.CustomerID, connstr);

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
        public static bool AddOrderInDB(Order obj, string connstr)
        {
            return InsertSQLQuery(obj, "Orders", connstr);
        }
        #endregion

        #region OrderUpdate
        public static bool UpdateOrderInDB(Order obj, string connstr)
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

            return UpdateSQLQuery(obj, "Orders", keys, connstr);
        }
        #endregion

        #region SelectCustomer
        public static List<Customer> GetCustomerFromDB(string customerNo, string customerType, string connstr)
        {
            List<Customer> returnValue = new List<Customer>();
            Dictionary<string, object> keys = new Dictionary<string, object>
            {
                { "ExternalId", customerNo },
                { "Type", customerType }
            };

            List<Customer> ret = ExecSQLQuery<Customer>(SelectCustomerByExternalIDCommand, keys, connstr);
            
            if(ret.Count > 0)
            {
                returnValue = ret;
            }

            return returnValue;

        }

        public static Customer GetCustomerFromDB(Guid ID, string connstr)
        {
            Customer returnValue = null;
            Dictionary<string, object> keys = new Dictionary<string, object>();
            keys.Add("CustomerID", ID);

            List<Customer> ret = ExecSQLQuery<Customer>(SelectCustomerByIDCommand, keys, connstr);

            if (ret.Count > 0)
            {
                returnValue = ret[0];
            }

            return returnValue;

        }
        #endregion

        #region CustomerUpdate
        public static bool UpdateCustomerInDB(Customer obj, string connstr)
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


            return UpdateSQLQuery(obj, "Customers", keys, connstr);
        }
        #endregion

        #region CustomerAdd
        public static bool AddCustomerInDB(Customer obj, string connstr)
        {
            return InsertSQLQuery(obj, "Customers", connstr);
        }
        #endregion

        #region Supporting
        public static bool Log(string message)
        {
            int affectedRows = 0;

            using (SqlConnection conn = new SqlConnection(SqlConnectionString))
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

        public static string GetUpdateSQLQuery<T>(T src, string tablename, Dictionary<string, object> keys)
        {
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
            query += " WHERE ";

            //Add conditions for update
            foreach (var kvp in keys)
            {
                Type valueType = kvp.Value.GetType();

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
                    else if(kvp.Value is DateTime)
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
        public static bool UpdateSQLQuery<T>(T src, string tablename, Dictionary<string, object> keys, string connstr)
        {
            int affectedRows = 0;
            string query = GetUpdateSQLQuery(src, tablename, keys);

            //Set all mapped properties of the update query
            var properties = src.GetType().GetProperties();
            List<PropertyInfo> mappedProperties = new List<PropertyInfo>(); 
            mappedProperties.AddRange(properties);

            using (SqlConnection conn = new SqlConnection(connstr))
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
                        foreach(var kvp in keys)
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
                }
            }

            return (affectedRows > 0);
        }

        /// <summary>
        /// Get the sql insert query as a string for an object. (Intended for internal use)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="src"></param>
        /// <param name="tablename"></param>
        /// <param name="keys"></param>
        /// <returns></returns>
        public static string GetSQLInsertQuery<T>(T src, string tablename)
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
                        query += value.ToString().Replace(",", ".");
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
                    else if (value is DateTime)
                    {
                        query += "'" + ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss") + "'";
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
        public static bool InsertSQLQuery<T>(T src, string tablename, string connstr)
        {
            int affectedRows = 0;
            string query = GetSQLInsertQuery(src, tablename);

            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlCommand command = new SqlCommand())
                {
                    conn.Open();
                    command.CommandText = query;
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    affectedRows = command.ExecuteNonQuery();
                    conn.Close();
                }
            }

            return (affectedRows > 0);
        }

        public static bool InsertSQLQuery(string query, string connstr)
        {
            int affectedRows = 0;

            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlCommand command = new SqlCommand())
                {
                    conn.Open();
                    command.CommandText = query;
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;
                    affectedRows = command.ExecuteNonQuery();
                    conn.Close();
                }
            }

            return (affectedRows > 0);
        }

        /// <summary>
        /// Fetch a list of matching objects from the database.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <param name="keys"></param>
        /// <param name="_config"></param>
        /// <returns></returns>
        public static List<T> ExecSQLQuery<T>(string query, Dictionary<string, object> keys, string connstr)
        {
            List<T> list = new List<T>();

            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlCommand command = new SqlCommand())
                {
                    conn.Open();
                    command.CommandText = query;
                    command.CommandType = System.Data.CommandType.Text;
                    command.Connection = conn;

                    foreach(var kvp in keys)
                    {
                        command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }

                    using (var result = command.ExecuteReader())
                    {
                        T obj = default(T);

                        while (result.Read())
                        {
                            obj = Activator.CreateInstance<T>();

                            foreach (PropertyInfo prop in obj.GetType().GetProperties())
                            {
                                if(result.GetColumnSchema().Any(c => c.ColumnName == prop.Name))
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

                    conn.Close();
                }
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
        public static bool ExecSQLNonQuery(string query, Dictionary<string, object> keys, string connstr)
        {
            int affectedRows = 0;

            using (SqlConnection conn = new SqlConnection(connstr))
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
                }
            }

            return (affectedRows > 0);
        }
        #endregion
    }
}
