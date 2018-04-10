using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RDotNet;
using System.Data;
using System.Reflection;

namespace SKIIBR
{
    public class RHandler
    {
        private string dataFilePath;
        public RHandler(string dataFilePath)
        {
            this.dataFilePath = dataFilePath;
        }

        public DataTable GetComplaintsData()
        {
            DataTable ParsedData = new DataTable();

            ParsedData.Columns.Add("Aktör");
            ParsedData.Columns.Add("Värde");

            REngine engine = REngine.GetInstance();
            engine.Evaluate("library(foreign)");
            dataFilePath = dataFilePath.Replace("\\", "\\\\");
            string evalExpression = String.Format("testdata<-read.spss('" + dataFilePath + "', to.data.frame=TRUE, trim.factor.names=TRUE, use.missings=TRUE, trim_values=TRUE)");
            var result = engine.Evaluate(evalExpression).AsDataFrame();

            var dataRows = ((IEnumerable<dynamic>)result.GetRows()).ToList();
            foreach (var dataRow in dataRows)
            {
                try
                {
                    ParsedData.Rows.Add(dataRow.Q1, dataRow.Q17);
                }
                catch
                { }
            }

            return ParsedData;
        }

        public DataTable GetSegmentData()
        {
            DataTable SegmentData = new DataTable();

            REngine engine = REngine.GetInstance();
            engine.Evaluate("library(foreign)");
            dataFilePath = dataFilePath.Replace("\\", "\\\\");
            string evalExpression = String.Format("testdata<-read.spss('" + dataFilePath + "', to.data.frame=TRUE, trim.factor.names=TRUE, use.missings=TRUE, trim_values=TRUE)");
            var result = engine.Evaluate(evalExpression).AsDataFrame();

            var dataRows = ((IEnumerable<dynamic>)result.GetRows()).ToList();

            List<string> cols = new List<string>();

            foreach (var prop in dataRows.First().DataFrame.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public))
            {
                var x = prop.Name;
                if (x == "ColumnNames")
                {
                    cols.AddRange(prop.GetValue(dataRows.First().DataFrame, null));
                }
            }
            SegmentData.Columns.Add("Aktör");

            foreach (var colName in cols)
            {
                SegmentData.Columns.Add(colName);
            }

            return SegmentData;

        }
            
    }
}
