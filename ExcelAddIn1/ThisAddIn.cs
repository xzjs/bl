using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {


            var bankAccounts = new List<Account> 
            {
                new Account 
                {
                    ID = 345,
                    Balance = 541.27
                },
                new Account 
                {
                    ID = 123,
                    Balance = -130.44
                }
            };

            DisplayInExcel(bankAccounts, (account, cell) =>
            // This multiline lambda expression sets custom processing rules  
            // for the bankAccounts.
            {
                cell.Value = account.ID;
                cell.Offset[0, 1].Value = account.Balance;
                if (account.Balance < 0)
                {
                    cell.Interior.Color = 255;
                    cell.Offset[0, 1].Interior.Color = 255;
                }
            });
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void DisplayInExcel(IEnumerable<Account> accounts,Action<Account, Excel.Range> DisplayFunc)
        {
            var excelApp = this.Application;
            // Add a new Excel workbook.
            excelApp.Workbooks.Add();
            excelApp.Visible = true;
            excelApp.Range["A1"].Value = "ID";
            excelApp.Range["B1"].Value = "Balance";
            excelApp.Range["A2"].Select();

            foreach (var ac in accounts)
            {
                DisplayFunc(ac, excelApp.ActiveCell);
                excelApp.ActiveCell.Offset[1, 0].Select();
            }
            // Copy the results to the Clipboard.
            excelApp.Range["A1:B3"].Copy();

            //根据内容自动调整列宽
            excelApp.Columns[1].AutoFit();
            excelApp.Columns[2].AutoFit();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
