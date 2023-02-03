using System;
using System.Drawing;
using SharpShell.SharpThumbnailHandler;
using System.IO;
using System.IO.Compression;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Text;
using System.Runtime.InteropServices;
using SharpShell.Attributes;

namespace ProcreateThumbnailExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.FileExtension, ".procreate")]
    public class ProcreateThumbnailHandler : SharpThumbnailHandler
    {
        protected override Bitmap GetThumbnailImage(uint width)
        {
            //  Create a stream reader for the selected item stream
            try
            {
                using (var reader = new ZipArchive(SelectedItemStream, ZipArchiveMode.Read))
                {
                    //  Now return a preview from the gcode or error when none present
                    return GetThumbnailFromProcreate(reader, width);
                }
            }
            catch (Exception exception)
            {
                //  Log the exception and return null for failure (no thumbnail to show)
                LogError("An exception occurred opening the file.", exception);
                return null;
            }
        }

        /// <summary>
        /// Create the resized bitmap representation that can be used by the Shell
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="width"></param>
        /// <returns></returns>
        private Bitmap GetThumbnailFromProcreate(ZipArchive reader, uint width)
        {
            //  Create the bitmap dimensions
            var thumbnailSize = new Size((int)width, (int)width);

            //  Create the bitmap
            var bitmap = new Bitmap(thumbnailSize.Width, thumbnailSize.Height,
                                    PixelFormat.Format32bppArgb);

            // Get pre-generated thumbnail from gcode file or error if none found
            Image procreateThumbnail = ReadThumbnailFromProcreate(reader);

            //  Create a graphics object to render to the bitmap
            using (var graphics = Graphics.FromImage(bitmap))
            {
                //  Set the rendering up for anti-aliasing
                graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                graphics.DrawImage(procreateThumbnail, 0, 0, thumbnailSize.Width, thumbnailSize.Height);
            }

            //  Return the bitmap
            return bitmap;
        }

        /// <summary>
        /// Parse the gcode recursively to find the last thumbnail and return the decoded image
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        private Image ReadThumbnailFromProcreate(ZipArchive reader)
        {
            var thumbnailEntry = reader.GetEntry("QuickLook/Thumbnail.png");
            if (thumbnailEntry == null) return null;

            // if no later thumbnail was found or its decoding ended with an error
            // we load the current one to image instance
            using (var ms = thumbnailEntry.Open())
            {
                Image image = Image.FromStream(ms, true);
                return image;
            }
        }
    }
}
