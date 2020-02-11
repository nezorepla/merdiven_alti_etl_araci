using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Text.RegularExpressions; 
using System.Diagnostics;

using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;

namespace adaptor
{
    class Program
    {

//################################################################################### burasi close button icin
  private const int MF_BYCOMMAND = 0x00000000;

    public const int SC_CLOSE = 0xF060;

    [DllImport("user32.dll")]
    public static extern int DeleteMenu(IntPtr hMenu, int nPosition, int wFlags);

    [DllImport("user32.dll")]
    private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

    [DllImport("kernel32.dll", ExactSpelling = true)]
    private static extern IntPtr GetConsoleWindow();
//################################################################################### burasi close button icin
 
 
 
//################################################################################### burasi system tray icin

		[DllImport("user32.dll", SetLastError = true)]
		static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

		[DllImport("user32.dll")]
		static extern IntPtr GetShellWindow();

		[DllImport("user32.dll")]
		static extern IntPtr GetDesktopWindow(); 
			

		static NotifyIcon notifyIcon;
		static IntPtr processHandle;
		static IntPtr WinShell;
		static IntPtr WinDesktop;
		static MenuItem HideMenu;
		static MenuItem RestoreMenu;


		static void Run()
		{
			Console.WriteLine("Listening to messages");

			while (true)
			{
		
				MainIslemler();
				Console.Read();
				
			}
		}


		private static void CleanExit(object sender, EventArgs e)
		{
			notifyIcon.Visible = false;
			Application.Exit();
			Environment.Exit(1);
		}
		

		static void Minimize_Click(object sender, EventArgs e)
		{			
			ResizeWindow(false);
		}


		static void Maximize_Click(object sender, EventArgs e)
		{
			ResizeWindow();
		}

		static void ResizeWindow(bool Restore = true)
		{
			if (Restore)
			{
				RestoreMenu.Enabled = false;
				HideMenu.Enabled = true;				
				SetParent(processHandle, WinDesktop);
			}
			else
			{
				RestoreMenu.Enabled = true;
				HideMenu.Enabled = false;
				SetParent(processHandle, WinShell);
			}
		}
//###################################################################################		
		
 
        public static SqlConnection baglanti;

        public static SqlConnection Tst_con;
        public static void exec(string q,int hata_bas)
        {
            SqlCommand cmd = new SqlCommand(q, baglanti);
cmd.CommandTimeout = 300000;  

            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
   try
            {
            cmd.ExecuteNonQuery();
            //      baglanti.Close();
    }
            catch (Exception ex)
            {
				if(hata_bas>0){
           Console.WriteLine("Hata");
		   Console.WriteLine(ex.ToString());
           Console.WriteLine("-------");
           Console.WriteLine(q);
		//   Console.ReadLine();
				}
            }
			}

        public static DataTable Getdata(string query)
        {
            // SqlConnection conn = new SqlConnection(GetConnStr(pConnKey));
            SqlDataAdapter da;
            DataTable dt = new DataTable();
            try
            {
                da = new SqlDataAdapter(query, baglanti);
                if (baglanti.State == ConnectionState.Closed)
                    baglanti.Open();
                da.Fill(dt);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {
                //   conn.Close();
            }
            return dt;
        }

public static string dtGlobal;
 public static string Fx_into(	string query) {

string q= query.ToUpper();
q=q.Replace("  ", " ");
q=q.Replace(System.Environment.NewLine," ");
q=q.Replace("  ", " ");
q=q.Trim();
q=q.Replace("FROM", " FROM");
q=q.Replace("DROP TABLE ", "DROP_TABLE_");
//q=q.Replace("USE ", "USE_");
q=q.Replace(System.Environment.NewLine," ");
q=q.Replace("DROP_TABLE_","INTO DROP_TABLE_");
//q=q.Replace("USE_","INTO U");
string rv= Left(vericek(q, "INTO ", " ").Trim(),40);

if(q.IndexOf("DROP ")>0)
{rv=q.Replace("INTO ", "");}
	
	
	return rv;
}
    public static string vericek(string StrData, string StrBas, string StrSon)
        {

            try
            {

                int IntBas = StrData.IndexOf(StrBas) + StrBas.Length;

                int IntSon = StrData.IndexOf(StrSon, IntBas + 1);

                return StrData.Substring(IntBas, IntSon - IntBas);

            }

            catch
            {

                return "";

            }

        }
	private static void Closer(string ToKill)
		{
			Process[] _proceses = null;
			_proceses = Process.GetProcessesByName(ToKill);
			foreach (Process proces in _proceses)
			{
				proces.Kill();
			}
		}
public  static TimeSpan lmt,bas;
 static void Main(string[] args)
        {
           
        Console.BackgroundColor = ConsoleColor.Blue;
        Console.Clear();
        Console.ForegroundColor = ConsoleColor.White;     
	

        DeleteMenu(GetSystemMenu(GetConsoleWindow(), false),SC_CLOSE, MF_BYCOMMAND);
		
		
		baglanti = new SqlConnection("Data Source=***********; Initial Catalog=**********; Integrated Security=true");
       Tst_con = new SqlConnection("Data Source=************; Initial Catalog=***********; Integrated Security=true");

//	  string filepath = "d://ConvertedFile.csv";
//      DataTable res = ConvertCSVtoDataTable(filepath);
	  
	  
var dateAndTime_Glb = DateTime.Now;   
dtGlobal =dateAndTime_Glb.ToString();
	   
 Console.WriteLine(dtGlobal);

  lmt = TimeSpan.Parse("10:00");   //10
  bas = dateAndTime_Glb.TimeOfDay; // 15

try{

notifyIcon = new NotifyIcon();	 	   
notifyIcon.Icon = new Icon("datamart.ico");
			notifyIcon.Text = "Datamart Monitor";			
			notifyIcon.Visible = true;

			ContextMenu menu = new ContextMenu();
			HideMenu = new MenuItem("Hide", new EventHandler(Minimize_Click));
			RestoreMenu = new MenuItem("Restore", new EventHandler(Maximize_Click));

			menu.MenuItems.Add(RestoreMenu);
			menu.MenuItems.Add(HideMenu);
			menu.MenuItems.Add(new MenuItem("Exit", new EventHandler(CleanExit)));

			notifyIcon.ContextMenu = menu;

  			Task.Factory.StartNew(Run);

			processHandle = Process.GetCurrentProcess().MainWindowHandle;
			
			WinShell = GetShellWindow();

			WinDesktop = GetDesktopWindow();
 		ResizeWindow(false);

        		Application.Run();
}
catch(Exception e1) {

Console.WriteLine("baslangic:");	

Console.WriteLine(e1.ToString());	


MainIslemler();
 
}


	   
Console.WriteLine("baslangic:");	
Console.WriteLine(dtGlobal);	

Console.WriteLine("bitis:");	
Console.WriteLine(DateTime.Now.ToString());	

System.Threading.Thread.Sleep(-1);
   Console.Read();
	}
		

		
    public static string Left(string gelen, int maxLength)
    {
        if (string.IsNullOrEmpty(gelen)) return gelen;
        maxLength = Math.Abs(maxLength);

        return ( gelen.Length <= maxLength 
               ? gelen 
               : gelen.Substring(0, maxLength)
               );
    }
	
	public static void MainIslemler(){



 
execFile("ALL_DROPS");

 if(!guncellik_kontrol("xxxxxxxxx")){

islem("xxxxxxxxxxxxx",0);    
execFile("xxxxxxxxxxxxx_DROP");
Doldur("xxxxxxxxxxxxxx");
 }	
 else { 
TempDoldur("xxxxxxxxxxxxx");
 }

 if(!guncellik_kontrol("yyyyyyyyyyyy")){
islem("yyyyyyyyyyyy",0);
execFile("yyyyyyyyyyyy_DROP");
Doldur("yyyyyyyyyyyy");
Doldur("ccccccccccccccc");

 }	
 else { 
TempDoldur("yyyyyyyyyyyy");
TempDoldur("ccccccccccccccc");
 }
 
 
 if(!guncellik_kontrol("dddddddddddddddddddd")){

islem("dddddddddddddddddddd",0);
execFile("dddddddddddddddddddd_DROP");
Doldur("dddddddddddddddddddd");
 }	
 else { 
TempDoldur("dddddddddddddddddddd");
 }
  


 
try{
		if (bas <= lmt)
		{
Process.Start(@"C:\Users\_MAIL.vbs");
		} /**/
}
catch(Exception ezz) {
 
Console.WriteLine(ezz.ToString());
//Console.ReadLine();
}


Closer("222222222222222");
Process.Start(@"C:\Users\222222222222222.exe");

 
	
	
System.Threading.Thread.Sleep(-1);

	}
public static void execFile(string isim){ 

Console.WriteLine(isim+" Basladi-->"+DateTime.Now.ToString());
string s_p=@"C:\Users\sql\jobs\"+isim+".sql";

string text;
var fileStream = new FileStream(s_p, FileMode.Open, FileAccess.Read);
using (var streamReader = new StreamReader(fileStream, Encoding.Default))
{
	text = streamReader.ReadToEnd();
}

			exec(text,0);


Console.WriteLine(isim+" Bitti-->"+DateTime.Now.ToString());
Console.WriteLine("##############################################################################");

}
	
public static void islem(string isim, int csv){

Console.WriteLine(isim+" Basladi-->"+DateTime.Now.ToString());


try{


string s_p=@"C:\Users\sql\jobs\"+isim+".sql";

string text;
var fileStream = new FileStream(s_p, FileMode.Open, FileAccess.Read);
using (var streamReader = new StreamReader(fileStream, Encoding.Default))
{
	text = streamReader.ReadToEnd();
}


string[] ayir = text.Split(';');

int  payda=	ayir.Length;
int pay=1;

foreach (string parca in ayir)
{	
int hata_bas= 1;
string fxi=	Fx_into(parca);
if(Left(fxi.Replace(" ", ""),4) == "DROP"||fxi.IndexOf("DROP")>0) { 
hata_bas=0;
}	

int pr=payda.ToString().Length;
			 
Console.WriteLine(pay.ToString().PadRight(pr) +"/"+ payda.ToString().PadRight(pr) +"---> "+fxi.PadRight(40)+" | "+Left(parca.Replace(System.Environment.NewLine,"").Replace("\r","").Replace("\n","").Replace("\t"," ").Trim(),40));


string araStr =  parca.Replace(System.Environment.NewLine,"");
		araStr=araStr.Replace(" ", "");
		araStr=araStr.Replace("\t", "");
		araStr=araStr.Replace("\n", "");
		araStr=araStr.Replace("\r", "");
		araStr=araStr.Trim();

if(araStr.Length>10){ 
			exec(parca,hata_bas);
			}
	 	
	 
		pay++;
 	 }			  



	





Console.WriteLine("##"+isim+" Tablosu tamamlandi"+DateTime.Now.ToString());


/* 
 */
if(csv==1){
DataTable dt = Getdata("SELECT * FROM ##"+isim);

//DataTableToCSV(dt,isim+".csv");
createXLS(dt,isim+".xls");
}

}
catch(Exception e) {

Console.WriteLine("main class");
Console.WriteLine(e.ToString());
//Console.ReadLine();
}
}
		 


public static void DataTableToCSV(DataTable dataTable, string filePath){



string pth=@"\\sunucu\klasor\Output\";
 string yol= pth+filePath;
	try
{
   File.Delete(pth+filePath); 
}
catch
{ 
Console.WriteLine("Log: "+yol+" adresi degisti");
	long time = DateTime.Now.Ticks;
	yol=pth+time.ToString()+"_"+filePath;
}



    StringBuilder fileContent = new StringBuilder();

        foreach (var col in dataTable.Columns) 
        {
            fileContent.Append(col.ToString() + "|");
        }

        fileContent.Replace("|", System.Environment.NewLine, fileContent.Length - 1, 1);

        foreach (DataRow dr in dataTable.Rows) 
        {
            foreach (var column in dr.ItemArray) 
            {
                fileContent.Append("\"" + column.ToString() + "\"|");
            }

            fileContent.Replace("|", System.Environment.NewLine, fileContent.Length - 1, 1);
        }

       System.IO.File.WriteAllText(yol, fileContent.ToString(),Encoding.Default); 
//	    Encoding.GetEncoding("Windows-1254")
//		Encoding.GetEncoding("iso-8859-9")
//		Encoding.GetEncoding(1254);

	   
 Console.WriteLine("Success: " +yol );
 } 

public static void mail(){ 
 
		
	}		
	
		
	public static DataTable ConvertCSVtoDataTable(string strFilePath)
 {
            StreamReader sr = new StreamReader(strFilePath);
            string[] headers = sr.ReadLine().Split(','); 
            DataTable dt = new DataTable();
            foreach (string header in headers)
            {
                dt.Columns.Add(header);
            }
            while (!sr.EndOfStream)
            {
                string[] rows = Regex.Split(sr.ReadLine(), ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                DataRow dr = dt.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dr[i] = rows[i];
                }
                dt.Rows.Add(dr);
            }
            return dt;
 } 
	
	
	
	
	
	/* 
ALTER proc [dbo].[SP_is_data_UpToDate] ( @tblname varchar(500)) 
as 
 
BEGIN TRY

--declare @tblname varchar(500) ='PK_DASHBOARD'

declare @q varchar(MAX)
set @q= 'select  CASE WHEN max(AS_OF_DT) = (YEAR(GETDATE()-1) *10000) +  (MONTH(GETDATE()-1) *100)  + DAY(GETDATE()-1) THEN 1 ELSE 0 END DURUM
 from '+ @tblname

 EXEC (@q)

END TRY
BEGIN CATCH
SELECT 0 DURUM
END CATCH
*/

private static bool guncellik_kontrol(string Tname)
{
	bool rv=false;
  SqlDataAdapter da1;
DataTable dt1 = new DataTable();
da1 = new SqlDataAdapter("SP_is_data_UpToDate '"+Tname+"'", Tst_con);
da1.Fill(dt1);    
string chk= dt1.Rows[0][0].ToString();
Console.WriteLine("SP_is_data_UpToDate  "+Tname);
	
if(chk=="1") 
{
	rv=true;
}

return rv;	
	
	
}


private static void TempDoldur(string Tname)
{
Console.WriteLine("-------------------------------");
Console.WriteLine(DateTime.Now.ToString());	
 if (Tst_con.State == ConnectionState.Closed)
{
Tst_con.Open();
}
 
/*
ALTER proc [dbo].[SP_Tbl_QueryGenerator] ( @tblname varchar(500)) 
as 
 
DECLARE @table_name SYSNAME
SELECT @table_name = case when upper(left(@tblname,4)) ='DBO.' then @tblname else 'dbo.'+@tblname  end ;

DECLARE 
      @object_name SYSNAME
    , @object_id INT

SELECT 
  --    @object_name = '[' + s.name + '].[' + o.name + ']'
     @object_name = '##' + o.name  
    , @object_id = o.[object_id]
FROM sys.objects o WITH (NOWAIT)
JOIN sys.schemas s WITH (NOWAIT) ON o.[schema_id] = s.[schema_id]
WHERE s.name + '.' + o.name = @table_name
    AND o.[type] = 'U'
    AND o.is_ms_shipped = 0

DECLARE @SQL NVARCHAR(MAX) = ''

;WITH index_column AS 
(
    SELECT 
          ic.[object_id]
        , ic.index_id
        , ic.is_descending_key
        , ic.is_included_column
        , c.name
    FROM sys.index_columns ic WITH (NOWAIT)
    JOIN sys.columns c WITH (NOWAIT) ON ic.[object_id] = c.[object_id] AND ic.column_id = c.column_id
    WHERE ic.[object_id] = @object_id
),
fk_columns AS 
(
     SELECT 
          k.constraint_object_id
        , cname = c.name
        , rcname = rc.name
    FROM sys.foreign_key_columns k WITH (NOWAIT)
    JOIN sys.columns rc WITH (NOWAIT) ON rc.[object_id] = k.referenced_object_id AND rc.column_id = k.referenced_column_id 
    JOIN sys.columns c WITH (NOWAIT) ON c.[object_id] = k.parent_object_id AND c.column_id = k.parent_column_id
    WHERE k.parent_object_id = @object_id
)
SELECT @SQL = 'CREATE TABLE ' + @object_name + CHAR(13) + '(' + CHAR(13) + STUFF((
    SELECT CHAR(9) + ', [' + c.name + '] ' + 
        CASE WHEN c.is_computed = 1
            THEN 'AS ' + cc.[definition] 
            ELSE replace(UPPER(tp.name),'İ','I') + 
                CASE WHEN tp.name IN ('varchar', 'char', 'varbinary', 'binary', 'text')
                       THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length AS VARCHAR(5)) END + ')'
                     WHEN tp.name IN ('nvarchar', 'nchar', 'ntext')
                       THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length / 2 AS VARCHAR(5)) END + ')'
                     WHEN tp.name IN ('datetime2', 'time2', 'datetimeoffset') 
                       THEN '(' + CAST(c.scale AS VARCHAR(5)) + ')'
                     WHEN tp.name = 'decimal' 
                       THEN '(' + CAST(c.[precision] AS VARCHAR(5)) + ',' + CAST(c.scale AS VARCHAR(5)) + ')'
                    ELSE ''
                END +
              CASE WHEN c.collation_name IS NOT NULL THEN ' COLLATE ' + c.collation_name ELSE '' END +
                CASE WHEN c.is_nullable = 1 THEN ' NULL' ELSE ' NOT NULL' END +
                CASE WHEN dc.[definition] IS NOT NULL THEN ' DEFAULT' + dc.[definition] ELSE '' END + 
                CASE WHEN ic.is_identity = 1 THEN ' IDENTITY(' + CAST(ISNULL(ic.seed_value, '0') AS CHAR(1)) + ',' + CAST(ISNULL(ic.increment_value, '1') AS CHAR(1)) + ')' ELSE '' END 
        END + CHAR(13)
    FROM sys.columns c WITH (NOWAIT)
    JOIN sys.types tp WITH (NOWAIT) ON c.user_type_id = tp.user_type_id
    LEFT JOIN sys.computed_columns cc WITH (NOWAIT) ON c.[object_id] = cc.[object_id] AND c.column_id = cc.column_id
    LEFT JOIN sys.default_constraints dc WITH (NOWAIT) ON c.default_object_id != 0 AND c.[object_id] = dc.parent_object_id AND c.column_id = dc.parent_column_id
    LEFT JOIN sys.identity_columns ic WITH (NOWAIT) ON c.is_identity = 1 AND c.[object_id] = ic.[object_id] AND c.column_id = ic.column_id
    WHERE c.[object_id] = @object_id
    ORDER BY c.column_id
    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, CHAR(9) + ' ')
    + ISNULL((SELECT CHAR(9) + ', CONSTRAINT [' + k.name + '] PRIMARY KEY (' + 
                    (SELECT STUFF((
                         SELECT ', [' + c.name + '] ' + CASE WHEN ic.is_descending_key = 1 THEN 'DESC' ELSE 'ASC' END
                         FROM sys.index_columns ic WITH (NOWAIT)
                         JOIN sys.columns c WITH (NOWAIT) ON c.[object_id] = ic.[object_id] AND c.column_id = ic.column_id
                         WHERE ic.is_included_column = 0
                             AND ic.[object_id] = k.parent_object_id 
                             AND ic.index_id = k.unique_index_id     
                         FOR XML PATH(N''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, ''))
            + ')' + CHAR(13)
            FROM sys.key_constraints k WITH (NOWAIT)
            WHERE k.parent_object_id = @object_id 
                AND k.[type] = 'PK'), '') + ')'  + CHAR(13)
    + ISNULL((SELECT (
        SELECT CHAR(13) +
             'ALTER TABLE ' + @object_name + ' WITH' 
            + CASE WHEN fk.is_not_trusted = 1 
                THEN ' NOCHECK' 
                ELSE ' CHECK' 
              END + 
              ' ADD CONSTRAINT [' + fk.name  + '] FOREIGN KEY(' 
              + STUFF((
                SELECT ', [' + k.cname + ']'
                FROM fk_columns k
                WHERE k.constraint_object_id = fk.[object_id]
                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
               + ')' +
              ' REFERENCES [' + SCHEMA_NAME(ro.[schema_id]) + '].[' + ro.name + '] ('
              + STUFF((
                SELECT ', [' + k.rcname + ']'
                FROM fk_columns k
                WHERE k.constraint_object_id = fk.[object_id]
                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
               + ')'
            + CASE 
                WHEN fk.delete_referential_action = 1 THEN ' ON DELETE CASCADE' 
                WHEN fk.delete_referential_action = 2 THEN ' ON DELETE SET NULL'
                WHEN fk.delete_referential_action = 3 THEN ' ON DELETE SET DEFAULT' 
                ELSE '' 
              END
            + CASE 
                WHEN fk.update_referential_action = 1 THEN ' ON UPDATE CASCADE'
                WHEN fk.update_referential_action = 2 THEN ' ON UPDATE SET NULL'
                WHEN fk.update_referential_action = 3 THEN ' ON UPDATE SET DEFAULT'  
                ELSE '' 
              END 
            + CHAR(13) + 'ALTER TABLE ' + @object_name + ' CHECK CONSTRAINT [' + fk.name  + ']' + CHAR(13)
        FROM sys.foreign_keys fk WITH (NOWAIT)
        JOIN sys.objects ro WITH (NOWAIT) ON ro.[object_id] = fk.referenced_object_id
        WHERE fk.parent_object_id = @object_id
        FOR XML PATH(N''), TYPE).value('.', 'NVARCHAR(MAX)')), '')
    + ISNULL(((SELECT
         CHAR(13) + 'CREATE' + CASE WHEN i.is_unique = 1 THEN ' UNIQUE' ELSE '' END 
                + ' NONCLUSTERED INDEX [' + i.name + '] ON ' + @object_name + ' (' +
                STUFF((
                SELECT ', [' + c.name + ']' + CASE WHEN c.is_descending_key = 1 THEN ' DESC' ELSE ' ASC' END
                FROM index_column c
                WHERE c.is_included_column = 0
                    AND c.index_id = i.index_id
                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '') + ')'  
                + ISNULL(CHAR(13) + 'INCLUDE (' + 
                    STUFF((
                    SELECT ', [' + c.name + ']'
                    FROM index_column c
                    WHERE c.is_included_column = 1
                        AND c.index_id = i.index_id
                    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '') + ')', '')  + CHAR(13)
        FROM sys.indexes i WITH (NOWAIT)
        WHERE i.[object_id] = @object_id
            AND i.is_primary_key = 0
            AND i.[type] = 2
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)')
    ), '')

select  @SQL as Q
--EXEC sys.sp_executesql @SQL
*/
//generate table script
 SqlDataAdapter da1;
DataTable dt1 = new DataTable();
da1 = new SqlDataAdapter("SP_Tbl_QueryGenerator '"+Tname+"'", Tst_con);
da1.Fill(dt1);    
string createQuery= dt1.Rows[0][0].ToString();
Console.WriteLine("generate table script ##"+Tname);

//create temp table
SqlCommand cmd_tmp = new SqlCommand(createQuery, baglanti);
cmd_tmp.CommandTimeout = 30000;  
cmd_tmp.ExecuteNonQuery();
System.Threading.Thread.Sleep(2000);
Console.WriteLine("create ##"+Tname);



//fill dt
SqlDataAdapter da;
DataTable dt = new DataTable();
da = new SqlDataAdapter("select  * from "+Tname, Tst_con);
da.Fill(dt);      
Console.WriteLine("fetch "+Tname);

 //var transaction = baglanti.BeginTransaction();
 //using (var bulkCopy = new SqlBulkCopy(baglanti, SqlBulkCopyOptions.KeepIdentity, transaction))
 using (SqlBulkCopy bulkCopy =new SqlBulkCopy(baglanti))
  {
		  bulkCopy.BatchSize = 5000;
		  bulkCopy.NotifyAfter = 100000;
		  bulkCopy.SqlRowsCopied += (sender, eventArgs) => Console.WriteLine("Wrote " + eventArgs.RowsCopied + " records.");

	bulkCopy.DestinationTableName = "##"+Tname;
	bulkCopy.WriteToServer(dt);
}

 Console.WriteLine("bulkCopy test to prod-> "+Tname);
System.Threading.Thread.Sleep(5000);

}


   
private static void Doldur(string Tname)
{
//string Tname ="PK_HESAP_DEVIR";	
 if (Tst_con.State == ConnectionState.Closed)
{
Tst_con.Open();
}

SqlCommand cmd_t = new SqlCommand("truncate table "+Tname, Tst_con);
cmd_t.CommandTimeout = 30000;  
cmd_t.ExecuteNonQuery();
 Console.WriteLine("truncate table "+Tname);
 
System.Threading.Thread.Sleep(2000);


/* yeni tablo create edilecekse

string sql = "Create Table ##"+Tname+" (";
foreach (DataColumn column in dt.Columns)
{
    sql += "[" + column.ColumnName + "] " + "nvarchar(500)" + ",";
}
sql = sql.TrimEnd(new char[] { ',' }) + ")";
SqlCommand cmd2 = new SqlCommand(sql, baglanti);
cmd2.CommandTimeout = 30000;  

            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
SqlDataAdapter da = new SqlDataAdapter(cmd2);
cmd2.ExecuteNonQuery();
*/

            DataTable dt =   Getdata("select * from ##"+Tname);

 Console.WriteLine("select * from ##"+Tname);

using (SqlBulkCopy bulkCopy =new SqlBulkCopy(Tst_con , SqlBulkCopyOptions.TableLock  ))
{	 

		  bulkCopy.BatchSize = 5000;
		  bulkCopy.NotifyAfter = 100000;
		  bulkCopy.SqlRowsCopied += (sender, eventArgs) => Console.WriteLine("Wrote " + eventArgs.RowsCopied + " records.");
 	bulkCopy.DestinationTableName = Tname;
	bulkCopy.WriteToServer(dt);
}

 Console.WriteLine("bulkCopy completed");


}

	
	public static DataTable ConvertExcelToDataTable(string FileName)  
{  
string fpth=@"\\sunucu\klasor\Input\";

    DataTable dtResult = null;  
    int totalSheet = 0; //No of sheets on excel file  
    using(OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+fpth+"" + FileName + ".xls;Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))  
    {  
        objConn.Open();  
        OleDbCommand cmd = new OleDbCommand();  
        OleDbDataAdapter oleda = new OleDbDataAdapter();  
        DataSet ds = new DataSet();  
        DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);  
        string sheetName = string.Empty;  
        if (dt != null)  
        {  
            var tempDataTable = (from dataRow in dt.AsEnumerable()  
            where!dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")  
            select dataRow).CopyToDataTable();  
            dt = tempDataTable;  
            totalSheet = dt.Rows.Count;  
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();  
        }  
        cmd.Connection = objConn;  
        cmd.CommandType = CommandType.Text;  
        cmd.CommandText = "SELECT * FROM [" + sheetName + "]";  
         oleda = new OleDbDataAdapter(cmd);  
        oleda.Fill(ds, "excelData");  
        dtResult = ds.Tables["excelData"];  
        objConn.Close();  
        return dtResult; //Returning Dattable  
    }  
}  

public static void CreateTableFromDt(DataTable dt,string Tname){
	
	
				 
			
	//    SqlConnection con = connection string ;
//con.Open();
string sql = "Create Table ##"+Tname+" (";
foreach (DataColumn column in dt.Columns)
{
    sql += "[" + column.ColumnName + "] " + "nvarchar(500)" + ",";
}
sql = sql.TrimEnd(new char[] { ',' }) + ")";
SqlCommand cmd2 = new SqlCommand(sql, baglanti);
cmd2.CommandTimeout = 30000;  

            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
SqlDataAdapter da = new SqlDataAdapter(cmd2);
cmd2.ExecuteNonQuery();



using (SqlBulkCopy bulkCopy =new SqlBulkCopy(baglanti/*, SqlBulkCopyOptions.KeepIdentity*/))
{
		  
		  bulkCopy.BatchSize = 5000;
		  bulkCopy.NotifyAfter = 100000;
		  bulkCopy.SqlRowsCopied += (sender, eventArgs) => Console.WriteLine("Wrote " + eventArgs.RowsCopied + " records.");

	bulkCopy.DestinationTableName ="##"+Tname;
	bulkCopy.WriteToServer(dt);


}
					
				
				
   }

   
   
 public static void createXLS(DataTable dt, string filePath){



string pth=@"\\sunucu\klasor\Output\";
 string yol= pth+filePath;
	try
{
   File.Delete(pth+filePath); 
}
catch
{ 
Console.WriteLine("Log: "+yol+" adresi degisti");
	long time = DateTime.Now.Ticks;
	yol=pth+time.ToString()+"_"+filePath;
}



//,Encoding.Default


            FileStream stream = new FileStream(yol, FileMode.OpenOrCreate);
            ExcelWriter writer = new ExcelWriter(stream);
            writer.BeginWrite();
 


          for (int i = 0; i < dt.Columns.Count; i++)
            {
                string name = dt.Columns[i].ColumnName.ToString();
                writer.WriteCell(0, i, name);

            }


				 for (int r = 0; r < dt.Rows.Count; r++)
                    {
 //string TempStr = HeadStr;
 //Sb.Append("{");
  for (int c = 0; c < dt.Columns.Count; c++)
                        {
							      writer.WriteCell(r+1, c, dt.Rows[r][c].ToString());
 //  TempStr = TempStr.Replace("<br>", Environment.NewLine).Replace(Dt.Columns[j] + j.ToString() + "¾", Dt.Rows[r][c].ToString());
 }
 //Sb.Append(TempStr + "},");
 }
				
				
				
				
				
            writer.EndWrite();
            stream.Close();
			
			
 Console.WriteLine("Success: " +yol );
 }

        /// <summary>
        /// Produces Excel file without using Excel
        /// </summary>
        public class ExcelWriter
        {
            private Stream stream;
            private BinaryWriter writer;

            private ushort[] clBegin = { 0x0809, 8, 0, 0x10, 0, 0 };
            private ushort[] clEnd = { 0x0A, 00 };


            private void WriteUshortArray(ushort[] value)
            {
                for (int i = 0; i < value.Length; i++)
                    writer.Write(value[i]);
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
            /// </summary>
            /// <param name="stream">The stream.</param>
            public ExcelWriter(Stream stream)
            {
                this.stream = stream;
                writer = new BinaryWriter(stream);
            }

            /// <summary>
            /// Writes the text cell value.
            /// </summary>
            /// <param name="row">The row.</param>
            /// <param name="col">The col.</param>
            /// <param name="value">The string value.</param>
            public void WriteCell(int row, int col, string value)
            {
			//	Encoding iso = Encoding.GetEncoding("ISO-8859-1");
//				Encoding utf8 = Encoding.UTF8;
				
                ushort[] clData = { 0x0204, 0, 0, 0, 0, 0 };
                int iLen = value.Length;
             //   byte[] plainText = Encoding.ASCII.GetBytes(value);
                byte[] plainText =Encoding.GetEncoding("ISO-8859-1").GetBytes(value);
				
				//byte[] isoBytes = Encoding.Convert(utf8,iso,utfBytes);
				
                clData[1] = (ushort)(8 + iLen);
                clData[2] = (ushort)row;
                clData[3] = (ushort)col;
                clData[5] = (ushort)iLen;
                WriteUshortArray(clData);
                writer.Write(plainText);
            }

            /// <summary>
            /// Writes the integer cell value.
            /// </summary>
            /// <param name="row">The row number.</param>
            /// <param name="col">The column number.</param>
            /// <param name="value">The value.</param>
            public void WriteCell(int row, int col, int value)
            {
                ushort[] clData = { 0x027E, 10, 0, 0, 0 };
                clData[2] = (ushort)row;
                clData[3] = (ushort)col;
                WriteUshortArray(clData);
                int iValue = (value << 2) | 2;
                writer.Write(iValue);
            }

            /// <summary>
            /// Writes the double cell value.
            /// </summary>
            /// <param name="row">The row number.</param>
            /// <param name="col">The column number.</param>
            /// <param name="value">The value.</param>
            public void WriteCell(int row, int col, double value)
            {
                ushort[] clData = { 0x0203, 14, 0, 0, 0 };
                clData[2] = (ushort)row;
                clData[3] = (ushort)col;
                WriteUshortArray(clData);
                writer.Write(value);
            }

            /// <summary>
            /// Writes the empty cell.
            /// </summary>
            /// <param name="row">The row number.</param>
            /// <param name="col">The column number.</param>
            public void WriteCell(int row, int col)
            {
                ushort[] clData = { 0x0201, 6, 0, 0, 0x17 };
                clData[2] = (ushort)row;
                clData[3] = (ushort)col;
                WriteUshortArray(clData);
            }

            /// <summary>
            /// Must be called once for creating XLS file header
            /// </summary>
            public void BeginWrite()
            {
                WriteUshortArray(clBegin);
            }

            /// <summary>
            /// Ends the writing operation, but do not close the stream
            /// </summary>
            public void EndWrite()
            {
                WriteUshortArray(clEnd);
                writer.Flush();
            }
        }

       
    }
}
