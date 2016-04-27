using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEIKReport.Common
{
    public class ZipTools
    {
        public string SourceFolderPath { get; set; }

        public ZipTools(string sourceFolderPath)
        {
            SourceFolderPath = sourceFolderPath;
        }

        public void ZipFolder(string zipFilePath)
        {
            using (Package package = Package.Open(zipFilePath, System.IO.FileMode.Create))
            {
                DirectoryInfo dir = new DirectoryInfo(SourceFolderPath);
                ZipDirectory(dir, package);
            }
        }

        private void ZipDirectory(DirectoryInfo dir, Package package)
        {
            foreach (FileInfo fi in dir.GetFiles())
            {
                string relativePath = fi.FullName.Replace(SourceFolderPath, string.Empty);
                relativePath = relativePath.Replace("\\", "/");
                Uri partUriDocument = PackUriHelper.CreatePartUri(new Uri(relativePath, UriKind.Relative));

                //resourcePath="Resource\Image.jpg"
                //Uri partUriResource = PackUriHelper.CreatePartUri(new Uri(resourcePath, UriKind.Relative));

                PackagePart part = package.CreatePart(partUriDocument,
                    System.Net.Mime.MediaTypeNames.Application.Zip);
                using (FileStream fs = fi.OpenRead())
                {
                    CopyStream(fs, part.GetStream());
                }

            }

            foreach (DirectoryInfo subDi in dir.GetDirectories())
            {
                ZipDirectory(subDi, package);
            }
        }

        private void CopyStream(Stream source, Stream target)
        {
            const int bufSize = 0x1000;
            byte[] buf = new byte[bufSize];
            int bytesRead = 0;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
            {
                target.Write(buf, 0, bytesRead);
            }
        }
    }
}
