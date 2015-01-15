using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using ADOX;
using System.Data.OleDb;
using System.Collections;
using System.Text.RegularExpressions;

namespace TagsReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private String GetMid(String input, String s, String e)
        {
            int pos = input.IndexOf(s);
            if (pos == -1)
            {
                return "";
            }

            pos += s.Length;

            int pos_end = 0;
            if (e == "")
            {
                pos_end = input.Length;
            }
            else
            {
                pos_end = input.IndexOf(e, pos);
            }

            if (pos_end == -1)
            {
                return "";
            }

            return input.Substring(pos, pos_end - pos);
        }

        private String trimTag(String tag)
        {
            String newTag = tag.Replace("&quot;", "");
            newTag = newTag.Replace("&amp;", "&");
            return newTag.Trim();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String path = textBox2.Text;
            String db_path = path + "\\Data.mdb";
            if (!File.Exists(db_path))
            {
                ADOX.Catalog catalog = new Catalog();
                catalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+db_path+";Jet OLEDB:Engine Type=5");

                ADODB.Connection cn = new ADODB.Connection();

                cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db_path, null, null, -1);
                catalog.ActiveConnection = cn;

                ADOX.Table table = new ADOX.Table();
                table.Name = "Account";

                ADOX.Column column = new ADOX.Column();
                column.ParentCatalog = catalog;
                column.Name = "ID";
                column.Type = DataTypeEnum.adInteger;
                column.DefinedSize = 9;
                column.Properties["AutoIncrement"].Value = true;
                table.Columns.Append(column, DataTypeEnum.adInteger, 9);
                //table.Keys.Append("FirstTablePrimaryKey", KeyTypeEnum.adKeyPrimary, column, null, null);
                table.Columns.Append("host", DataTypeEnum.adVarWChar, 50);
                table.Columns.Append("code", DataTypeEnum.adVarWChar, 255);
                table.Columns.Append("refer", DataTypeEnum.adVarWChar, 50);
                table.Columns.Append("thread", DataTypeEnum.adInteger, 9);
                table.Columns.Append("company", DataTypeEnum.adVarWChar, 50);
                catalog.Tables.Append(table);

                cn.Close();
            }
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db_path);
            conn.Open();
            OleDbCommand sql = conn.CreateCommand();
            sql.CommandText = "DELETE FROM Account";
            sql.ExecuteNonQuery();

            DirectoryInfo dir = new DirectoryInfo(path);
            foreach(FileInfo file in dir.GetFiles())
            {
                if (file.Name.Contains(".html"))
                {
                    String html = file.OpenText().ReadToEnd();
                    String tags = GetMid(html, "<h1>AppNexus Console Placement Information</h1>", "");
                    String company = textBox1.Text;
                    if (tags.Length > 0)
                    {
                        String refer = "http://" + GetMid(tags, "Placement Group: <b>", "</b>");
                        //Regex urlRegex = new Regex(@"((http|https)://)?(www.)?[a-z0-9\.]+(\.(com|net|cn|com\.cn|com\.net|net\.cn))(/[^\s\n]*)?");
                        //Match match = urlRegex.Match(refer);
                        //refer = match.Value;
                        String oneTag = GetMid(tags, "SCRIPT SRC=", "TYPE=");
                        tags = GetMid(tags, oneTag, "");
                        oneTag = trimTag(oneTag);
                        while(!String.IsNullOrEmpty(oneTag))
                        {
                            String host = GetMid(oneTag, "http://", "/");
                            String code = GetMid(oneTag, host, "");
                            sql.CommandText = "INSERT INTO Account(host,code,refer,thread,company) VALUES ('" + host + "','" + code + "','" + refer + "',1,'" + company + "')";
                            sql.ExecuteNonQuery();
                            oneTag = GetMid(tags, "SCRIPT SRC=", "TYPE=");
                            tags = GetMid(tags, oneTag, "");
                            oneTag = trimTag(oneTag);
                        }
                    }
                }
                else if(file.Name.Contains(".txt"))
                {
                    String html = file.OpenText().ReadToEnd();
                    String tags = GetMid(html, "AppNexus Console Placement Information", "");
                    String company = textBox1.Text;
                    if (tags.Length > 0)
                    {
                        String refer = "http://" + GetMid(tags, "Placement Group: ", "\r\n");
                        //Regex urlRegex = new Regex(@"((http|https)://)?(www.)?[a-z0-9\.]+(\.(com|net|cn|com\.cn|com\.net|net\.cn))(/[^\s\n]*)?");
                        //Match match = urlRegex.Match(refer);
                        //refer = match.Value;
                        String oneTag = GetMid(tags, "SCRIPT SRC=", "TYPE=");
                        tags = GetMid(tags, oneTag, "");
                        oneTag = trimTag(oneTag);
                        while (!String.IsNullOrEmpty(oneTag))
                        {
                            String host = GetMid(oneTag, "http://", "/");
                            String code = GetMid(oneTag, host, "\"");
                            sql.CommandText = "INSERT INTO Account(host,code,refer,thread,company) VALUES ('" + host + "','" + code + "','" + refer + "',1,'" + company + "')";
                            sql.ExecuteNonQuery();
                            oneTag = GetMid(tags, "SCRIPT SRC=", "TYPE=");
                            tags = GetMid(tags, oneTag, "");
                            oneTag = trimTag(oneTag);
                        }
                    }
                }
            }
        }
    }
}
