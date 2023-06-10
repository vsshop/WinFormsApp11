using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.ApplicationServices;
using WinFormsApp11.Models;
using Excel =  Microsoft.Office.Interop.Excel.Application;

namespace WinFormsApp11
{
    internal class ExcelApplication
    {
        Excel application;
        Workbook book;
        Worksheet sheet;
        public ExcelApplication()
        {
            application = new Excel();

#if DEBUG
            application.Visible = true;
#else
            application.Visible = false;
#endif
        }

        public void InitWorkbook()
        {
            if(book != null)
            {
                book.Close(true);
            }
            book = application.Workbooks.Add();
            sheet = book.Sheets[1];
        }
        
        public void InitWorkbook(string path)
        {
            if (book != null)
            {
                book.Close(true);
            }
            book = application.Workbooks.Open(path);
            sheet = book.Sheets[1];
        }
        public void Exit()
        {
            if(book != null)
            {

                book.Close(true);
            }

            application.Quit();
        }

        public UserModel[] GetUsers()
        {
            if(sheet == null)
            {
                return new UserModel[0];
            }

            int count = Count;
            UserModel[] users = new UserModel[count];

            for (int i = 0; i < count; i++)
            {
                var newUser = new UserModel();
                newUser.Name = sheet.Cells[i + 1, "A"].Text;
                newUser.Role = sheet.Cells[i + 1, "B"].Text;
                newUser.System = sheet.Cells[i + 1, "C"].Text;
                newUser.IP = sheet.Cells[i + 1, "D"].Text;

                if(double.TryParse(sheet.Cells[i + 1, "E"].Text, out double result))
                {
                    newUser.Speed = result;
                }
                

                users[i] = newUser;
            };
            return users;
        }

        public void SetUsers(UserModel[] users)
        {
            if(sheet == null)
            {
                return;
            }

            for (int i = 0; i < users.Length; i++)
            {
                sheet.Cells[i + 1, "A"].Value2 = users[i].Name;
                sheet.Cells[i + 1, "B"].Value2 = users[i].Role;
                sheet.Cells[i + 1, "C"].Value2 = users[i].System;
                sheet.Cells[i + 1, "D"].Value2 = users[i].IP;
                sheet.Cells[i + 1, "E"].Value2 = users[i].Speed;
            }

            ClearSheet(users.Length);
        }

        void ClearSheet(int from)
        {
            int currentCount = from;
            while (true)
            {
                var value = sheet.Cells[currentCount + 1, "A"].Value2;
                if (value == null)
                {
                    break;
                }

                sheet.Cells[currentCount + 1, "A"].Value2 = null;
                sheet.Cells[currentCount + 1, "B"].Value2 = null;
                sheet.Cells[currentCount + 1, "C"].Value2 = null;
                sheet.Cells[currentCount + 1, "D"].Value2 = null;
                sheet.Cells[currentCount + 1, "E"].Value2 = null;
                currentCount++;
            }
        }

        public int Count
        {
            get
            {
                int count = 0;
                if(sheet == null)
                {
                    return -1;
                }

                while (true)
                {
                    var value = sheet.Cells[count + 1, "A"].Value2;
                    if(value == null)
                    {
                        break;
                    }
                    count++;
                }

                return count;
            }
        }
    }
}
