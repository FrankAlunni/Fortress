//-----------------------------------------------------------------------
// <copyright file="Engine.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Frank Alunni</author>
//-----------------------------------------------------------------------
namespace B1C.Utility.Imaging
{
    using System.Drawing;
    using System.Drawing.Drawing2D;
    using System.Drawing.Imaging;

    /// <summary>
    /// Instance of the Engine class.
    /// </summary>
    /// <author>Frank Alunni</author>
    public class Engine
    {
        /// <summary>
        /// Resizes an image to fit inside the given dimension
        /// </summary>
        /// <remarks>
        /// This will place borders of the given backround colour if the image doesn't fit 
        /// exactly in the given dimensions
        /// </remarks>
        /// <param name="originalImage">The image to resize</param>
        /// <param name="width">The new image width (in pixels)</param>
        /// <param name="height">The new image height (in pixels)</param>
        /// <param name="resolution">The resolution of the new image (in dpi)</param>
        /// <param name="backgroundColour">The backround colour</param>
        /// <returns>A resized image</returns>
        public static Image Resize(Image originalImage, int width, int height, float resolution, Color backgroundColour)
        {
            int destX = 0;
            int destY = 0;
            int destWidth = 0;
            int destHeight = 0;

            int sourceWidth = originalImage.Width;
            int sourceHeight = originalImage.Height;
            float widthPercentage = width / (float)sourceWidth;
            float heightPercentage = height / (float)sourceHeight;

            if (heightPercentage < widthPercentage)
            {
                destHeight = height;
                destWidth = (int)(heightPercentage * sourceWidth);
                destY = 0;
                destX = (width - destWidth) / 2;
            }
            else
            {
                destWidth = width;
                destHeight = (int)(widthPercentage * sourceHeight);
                destX = 0;
                destY = (height - destHeight) / 2;
            }                        

            var photoBitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);            
            photoBitmap.SetResolution(resolution, resolution);            

            Graphics photoGraphics = Graphics.FromImage(photoBitmap);

            // Set the background colour
            photoGraphics.Clear(backgroundColour); 
            photoGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic;

            photoGraphics.DrawImage(originalImage, new Rectangle(destX, destY, destWidth, destHeight), new Rectangle(0, 0, sourceWidth, sourceHeight), GraphicsUnit.Pixel);

            photoGraphics.Dispose();
            originalImage.Dispose();

            return photoBitmap;
        }

        /// <summary>
        /// Crops the specified img photo.
        /// </summary>
        /// <param name="originalImage">The original image.</param>
        /// <param name="width">The input width.</param>
        /// <param name="height">The input height.</param>
        /// <param name="resolution">The resolution.</param>
        /// <returns>Resized image</returns>
        public static Image Crop(Image originalImage, int width, int height, float resolution)
        {
            int destX = 0;
            int destY = 0;

            int sourceWidth = originalImage.Width;
            int sourceHeight = originalImage.Height;
            float widthPercentage = width / (float)sourceWidth;
            float heightPercentage = height / (float)sourceHeight;
            float percentValue = 0;

            if (heightPercentage < widthPercentage)
            {
                percentValue = widthPercentage;
                destY = (int)((height - (sourceHeight * percentValue)) / 2);
            }
            else
            {
                percentValue = heightPercentage;
                destX = (int)((width - (sourceWidth * percentValue)) / 2);
            }

            var destWidth = (int)(sourceWidth * percentValue);
            var destHeight = (int)(sourceHeight * percentValue);

            var photoBitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            photoBitmap.SetResolution(resolution, resolution);

            Graphics photoGraphics = Graphics.FromImage(photoBitmap);
            photoGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic;

            photoGraphics.DrawImage(originalImage, new Rectangle(destX, destY, destWidth, destHeight), new Rectangle(destX, destY, sourceWidth, sourceHeight), GraphicsUnit.Pixel);

            photoGraphics.Dispose();
            originalImage.Dispose();

            return photoBitmap;
        }
    }
}
