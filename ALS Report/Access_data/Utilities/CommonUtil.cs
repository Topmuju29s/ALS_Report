using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Access_data.Utilities
{
   public class CommonUtil
    {
        public static DataTable ConvertDoListToDataTable<T>(List<T> objDoList)
        {

            DataTable dtOut = new DataTable();

            if (objDoList != null && objDoList.Count >= 0)
            {
                T objSource = System.Activator.CreateInstance<T>();
                if (objDoList.Count > 0 && objDoList[0] != null)
                {
                    objSource = objDoList[0];
                }
                //Generate DataTable Column
                PropertyInfo[] pSourceInfo = objSource.GetType().GetProperties().Where(d => !d.PropertyType.FullName.Contains("System.Data")).ToArray();
                foreach (PropertyInfo pInfo in pSourceInfo)
                {
                    string strPropertyType = string.Empty;
                    if (pInfo.PropertyType.FullName == objSource.GetType().ToString())
                    {
                        continue;
                    }

                    if (pInfo.PropertyType.IsGenericType && pInfo.PropertyType.Name.Contains("Nullable"))
                    {
                        Type tNullableType = Type.GetType(pInfo.PropertyType.FullName);
                        strPropertyType = tNullableType.GetGenericArguments()[0].FullName;
                    }
                    else if (!pInfo.PropertyType.IsGenericType)
                    {
                        strPropertyType = pInfo.PropertyType.FullName;
                    }
                    else
                    {
                        continue;
                    }
                    DataColumn col = new DataColumn(pInfo.Name, Type.GetType(strPropertyType));
                    dtOut.Columns.Add(col);
                }

                // Transfer Data from Do list to DataTable
                foreach (T obj in objDoList)
                {
                    if (obj != null)
                    {
                        DataRow row = dtOut.NewRow();
                        for (int idx = 0; idx < dtOut.Columns.Count; idx++)
                        {
                            PropertyInfo pDestInfo = obj.GetType().GetProperty(dtOut.Columns[idx].ColumnName);
                            Object objVal = pDestInfo.GetValue(obj, null);
                            row[dtOut.Columns[idx].ColumnName] = objVal == null ? DBNull.Value : objVal;
                        }
                        dtOut.Rows.Add(row);
                        dtOut.AcceptChanges();
                    }
                }
            }
            return dtOut;
        }
    }
}
