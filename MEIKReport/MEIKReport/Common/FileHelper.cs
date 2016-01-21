﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;

namespace MEIKReport.Common
{
    public class FileHelper
    {
        public static string GetXPSFromDialog(bool isSaved)
        {
            if (isSaved)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();

                saveFileDialog.Filter = "XPS Document files (*.xps)|*.xps|Txt Document files(*.txt)|*.txt";
                saveFileDialog.FilterIndex = 1;

                if (saveFileDialog.ShowDialog() == true)
                {
                    return saveFileDialog.FileName;
                }
                else
                {
                    return null;
                }
            }
            else return string.Format("{0}\\temp.xps", Environment.CurrentDirectory);//制造一个临时存储
        }

        public static string OpenXPSFileFromDialog()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XPS Document Files(*.xps)|*.xps";

            if (openFileDialog.ShowDialog() == true)
                return openFileDialog.FileName;
            else
                return null;
        }
        public static bool SaveXPS(FixedPage page, bool isSaved)
        {
            FixedDocument fixedDoc = new FixedDocument();//创建一个文档
            fixedDoc.DocumentPaginator.PageSize = new Size(96 * 8.5, 96 * 11);

            PageContent pageContent = new PageContent();
            ((IAddChild)pageContent).AddChild(page);
            fixedDoc.Pages.Add(pageContent);//将对象加入到当前文档中

            string containerName = GetXPSFromDialog(isSaved);
            if (containerName != null)
            {
                try
                {
                    File.Delete(containerName);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }

                XpsDocument _xpsDocument = new XpsDocument(containerName, FileAccess.Write);

                XpsDocumentWriter xpsdw = XpsDocument.CreateXpsDocumentWriter(_xpsDocument);
                xpsdw.Write(fixedDoc);//写入XPS文件
                _xpsDocument.Close();
                return true;
            }
            else return false;
        }

        static XpsDocument xpsPackage = null;
        public static void LoadDocumentViewer(string xpsFileName, DocumentViewer viewer)
        {
            XpsDocument oldXpsPackage = xpsPackage;//保存原来的XPS包
            xpsPackage = new XpsDocument(xpsFileName, FileAccess.Read, CompressionOption.NotCompressed);//从文件中读取XPS文档

            FixedDocumentSequence fixedDocumentSequence = xpsPackage.GetFixedDocumentSequence();//从XPS文档对象得到FixedDocumentSequence
            viewer.Document = fixedDocumentSequence as IDocumentPaginatorSource;

            if (oldXpsPackage != null)
                oldXpsPackage.Close();
            xpsPackage.Close();
        }
    }
}