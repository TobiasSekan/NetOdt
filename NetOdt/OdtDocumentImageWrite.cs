using NetOdt.Constants;
using NetOdt.Helper;
using System;
using System.IO;

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument
    {
        /// <summary>
        /// Add a picture to the document
        /// </summary>
        /// <param name="imagePath"></param>
        /// <param name="width">The width of the picture in centimeter (cm)</param>
        /// <param name="height">The height of the picture in centimeter (cm)</param>
        public void Write(string imagePath, double width, double height)
        {
            PictureCount++;

            var pictureExtension = Path.GetExtension(imagePath);
            var mineType         = FileHelper.GetMineType(imagePath);
            var picturePath      = $"{FolderResource.PictureFolderName}/{PictureCount}{pictureExtension}";

            FileHelper.Copy(imagePath, UriHelper.Combine(TempWorkingUri, picturePath));

            TextContent.Append($"<draw:frame draw:style-name=\"fr1\" draw:name=\"Picture{PictureCount}\" text:anchor-type=\"paragraph\" svg:width=\"{width}cm\" svg:height=\"{height}cm\" draw:z-index=\"0\">");
            TextContent.Append($"<draw:image xlink:href=\"{picturePath}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\" loext:mime-type=\"{mineType}\"/>");
            TextContent.Append($"</draw:frame>");
        }

        /// <summary>
        /// Add a picture to the document
        /// </summary>
        /// <param name="imageUri"></param>
        /// <param name="width">The width of the picture in centimeter (cm)</param>
        /// <param name="height">The height of the picture in centimeter (cm)</param>
        public void Write(Uri imageUri, double width, double height)
        {
            var pictureName      = Guid.NewGuid();
            var pictureExtension = PathHelper.GetExtension(imageUri);
            var mineType         = FileHelper.GetMineType(imageUri.AbsolutePath);
            var picturePath      = $"{FolderResource.PictureFolderName}/{PictureCount}{pictureExtension}";

            FileHelper.Copy(imageUri, UriHelper.Combine(TempWorkingUri, picturePath));

            TextContent.Append($"<draw:frame draw:style-name=\"fr1\" draw:name=\"Picture{PictureCount}\" text:anchor-type=\"paragraph\" svg:width=\"{width}cm\" svg:height=\"{height}cm\" draw:z-index=\"0\">");
            TextContent.Append($"<draw:image xlink:href=\"{picturePath}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\" loext:mime-type=\"{mineType}\"/>");
            TextContent.Append($"</draw:frame>");
        }
    }
}
