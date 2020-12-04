using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Data.SQLite;
using System.Text;

namespace NTPEmailMarketing
{
    public class Delay
    {
        public const string TableName = "tb_Delay";     
        public int Value { get; set; }
        
        public static int GetValue()
        {
            string SQL = String.Format("SELECT Value FROM {0}", TableName);
            int value = 0;            
            using (DB db = new DB())
            {
                try
                {
                    CreateTable(db);
                    using (var reader = db.GetReader(SQL))
                    {
                        if (reader.HasRows)
                        {
                            if (reader.Read())
                            {
                                value = reader["Value"] == null ? 0 : (int)reader["Value"];
                            }
                        }       
                    }
                    return value;
                }
                catch(Exception ex)
                {
                    return 0;
                }
            }
        }
        public static bool CreateTable(DB db)
        {
            bool ret = true;
            //using (DB db = new DB())
            //{
                //db.CreateTry();
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("create table if not exists {0}(Value INTEGER NULL); ", TableName);
                ret = db.UpdateTry(sb.ToString());
            //}
            return ret;
        }
        public static bool Insert(int value)
        {
            bool ret = false;                        
            using (DB db = new DB())
            {
                CreateTable(db);
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("INSERT INTO {0}(Value) VALUES (@Value); ", TableName);
                SQLiteParameter[] parameters = new SQLiteParameter[1];
                parameters[0] = new SQLiteParameter("@Value", value);
                ret = db.UpdateTry(sb.ToString(), parameters);
            }                        
            return ret;
        }
        public static bool Update(int Value)
        {
            bool ret = false;
            using (DB db = new DB())
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("UPDATE {0} SET Value=@Value", TableName);
                SQLiteParameter[] parameters = new SQLiteParameter[1];
                parameters[0] = new SQLiteParameter("@Value", Value);
                ret = db.UpdateTry(sb.ToString(), parameters);
            }
            return ret;
        }
        public static bool Delete()
        {
            long ret = -1;
            using (DB db = new DB())
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("DELETE from {0}", TableName);
                ret = db.ExecuteTry(sb.ToString(), null);
            }
            return ret > -1;
        }
    }
}
