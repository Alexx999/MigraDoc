using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering.Resources;
using PdfSharp.Drawing;
using PdfSharp.Xps.Rendering;
using PdfSharp.Xps.XpsModel;

namespace MigraDoc.Rendering
{
    class XpsRenderer : ShapeRenderer
    {
        private readonly Image _image;
        private readonly Dictionary<string, XpsDocument> _xpsCache;
        private ImageFailure _failure;
        private string _imageFilePath;
        private int _index;

        internal XpsRenderer(XGraphics gfx, Image image, FieldInfos fieldInfos,
            Dictionary<string, XpsDocument> xpsCache)
            : base(gfx, image, fieldInfos)
        {
            _image = image;
            _xpsCache = xpsCache;
            ImageRenderInfo renderInfo = new ImageRenderInfo();
            renderInfo.DocumentObject = _shape;
            _renderInfo = renderInfo;
        }

        internal XpsRenderer(XGraphics gfx, RenderInfo renderInfo, FieldInfos fieldInfos,
            Dictionary<string, XpsDocument> xpsCache)
            : base(gfx, renderInfo, fieldInfos)
        {
            _xpsCache = xpsCache;
            _image = (Image)renderInfo.DocumentObject;
        }

        internal string GetFileName(string fullPath, out int index)
        {
            index = 0;
            var split = fullPath.LastIndexOf(".xps:", StringComparison.Ordinal);
            if (split != -1)
            {
                var indexStr = fullPath.Substring(split + 5);
                int.TryParse(indexStr, out index);
                return fullPath.Substring(0, split + 4);
            }

            return fullPath;
        }

        internal override void Format(Area area, FormatInfo previousFormatInfo)
        {
            var fullPath = _image.GetFilePath(_documentRenderer.WorkingDirectory);
            _imageFilePath = GetFileName(fullPath, out _index);
            // The Image is stored in the string if path starts with "base64:", otherwise we check whether the file exists.
            if (!_imageFilePath.StartsWith("base64:") &&
                !File.Exists(_imageFilePath))
            {
                _failure = ImageFailure.FileNotFound;
                Debug.WriteLine(Messages2.ImageNotFound(_image.Name), "warning");
            }
            ImageFormatInfo formatInfo = (ImageFormatInfo)_renderInfo.FormatInfo;
            formatInfo.Failure = _failure;
            formatInfo.ImagePath = fullPath;
            CalculateImageDimensions();
            base.Format(area, previousFormatInfo);
        }

        protected override XUnit ShapeHeight
        {
            get
            {
                ImageFormatInfo formatInfo = (ImageFormatInfo)_renderInfo.FormatInfo;
                return formatInfo.Height + _lineFormatRenderer.GetWidth();
            }
        }

        protected override XUnit ShapeWidth
        {
            get
            {
                ImageFormatInfo formatInfo = (ImageFormatInfo)_renderInfo.FormatInfo;
                return formatInfo.Width + _lineFormatRenderer.GetWidth();
            }
        }

        internal override void Render()
        {
            RenderFilling();

            ImageFormatInfo formatInfo = (ImageFormatInfo)_renderInfo.FormatInfo;
            Area contentArea = _renderInfo.LayoutInfo.ContentArea;
            XRect destRect = new XRect(contentArea.X, contentArea.Y, formatInfo.Width, formatInfo.Height);

            if (formatInfo.Failure == ImageFailure.None)
            {
                FixedPage xImage = null;
                try
                {
                    XRect srcRect = new XRect(formatInfo.CropX, formatInfo.CropY, formatInfo.CropWidth, formatInfo.CropHeight);
                    //xImage = XImage.FromFile(formatInfo.ImagePath);
                    xImage = GetFixedPage(formatInfo.ImagePath);
                    RenderXps(xImage, destRect, srcRect, XGraphicsUnit.Point); //Pixel.
                }
                catch (Exception)
                {
                    RenderFailureImage(destRect);
                }
            }
            else
                RenderFailureImage(destRect);

            RenderLine();
        }

        private void RenderXps(FixedPage xImage, XRect destRect, XRect srcRect, XGraphicsUnit point)
        {
            var page = _gfx.PdfPage;
            var context = new DocumentRenderingContext(page.Owner);

            using (XForm form = new XForm(page.Owner, XUnit.FromPoint(xImage.PointWidth), XUnit.FromPoint(xImage.PointHeight)))
            {

                var writer = new PdfContentWriter(context, form, RenderMode.Default);


                writer.BeginContent(false);
                writer.WriteElements(xImage.Content);
                writer.EndContent();

                _gfx.DrawImage(form, destRect, srcRect, point);
            }
        }

        void RenderFailureImage(XRect destRect)
        {
            _gfx.DrawRectangle(XBrushes.LightGray, destRect);
            string failureString;
            ImageFormatInfo formatInfo = (ImageFormatInfo)RenderInfo.FormatInfo;

            switch (formatInfo.Failure)
            {
                case ImageFailure.EmptySize:
                    failureString = Messages2.DisplayEmptyImageSize;
                    break;

                case ImageFailure.FileNotFound:
                    failureString = Messages2.DisplayImageFileNotFound;
                    break;

                case ImageFailure.InvalidType:
                    failureString = Messages2.DisplayInvalidImageType;
                    break;

                case ImageFailure.NotRead:
                default:
                    failureString = Messages2.DisplayImageNotRead;
                    break;
            }

            // Create stub font
            XFont font = new XFont("Courier New", 8);
            _gfx.DrawString(failureString, font, XBrushes.Red, destRect, XStringFormats.Center);
        }

        private void CalculateImageDimensions()
        {
            ImageFormatInfo formatInfo = (ImageFormatInfo)_renderInfo.FormatInfo;

            if (formatInfo.Failure == ImageFailure.None)
            {
                FixedPage xImage = null;
                try
                {
                    //xImage = XImage.FromFile(_imageFilePath);
                    xImage = GetFixedPage(_imageFilePath, _index);
                }
                catch (InvalidOperationException ex)
                {
                    Debug.WriteLine(Messages2.InvalidImageType(ex.Message));
                    formatInfo.Failure = ImageFailure.InvalidType;
                }

                if (formatInfo.Failure == ImageFailure.None)
                {
                    try
                    {
                        XUnit usrWidth = _image.Width.Point;
                        XUnit usrHeight = _image.Height.Point;
                        bool usrWidthSet = !_image._width.IsNull;
                        bool usrHeightSet = !_image._height.IsNull;

                        XUnit resultWidth = usrWidth;
                        XUnit resultHeight = usrHeight;

                        Debug.Assert(xImage != null);
                        double xPixels = xImage.Width;
                        bool usrResolutionSet = !_image._resolution.IsNull;

                        double horzRes = usrResolutionSet ? _image.Resolution : 96;
                        double vertRes = usrResolutionSet ? _image.Resolution : 96;

// ReSharper disable CompareOfFloatsByEqualityOperator
                        if (horzRes == 0 && vertRes == 0)
                        {
                            horzRes = 72;
                            vertRes = 72;
                        }
                        else if (horzRes == 0)
                        {
                            Debug.Assert(false, "How can this be?");
                            horzRes = 72;
                        }
                        else if (vertRes == 0)
                        {
                            Debug.Assert(false, "How can this be?");
                            vertRes = 72;
                        }
                        // ReSharper restore CompareOfFloatsByEqualityOperator

                        XUnit inherentWidth = XUnit.FromInch(xPixels / horzRes);
                        double yPixels = xImage.Height;
                        XUnit inherentHeight = XUnit.FromInch(yPixels / vertRes);

                        //bool lockRatio = _image.IsNull("LockAspectRatio") ? true : _image.LockAspectRatio;
                        bool lockRatio = _image._lockAspectRatio.IsNull || _image.LockAspectRatio;

                        double scaleHeight = _image.ScaleHeight;
                        double scaleWidth = _image.ScaleWidth;
                        //bool scaleHeightSet = !_image.IsNull("ScaleHeight");
                        //bool scaleWidthSet = !_image.IsNull("ScaleWidth");
                        bool scaleHeightSet = !_image._scaleHeight.IsNull;
                        bool scaleWidthSet = !_image._scaleWidth.IsNull;

                        if (lockRatio && !(scaleHeightSet && scaleWidthSet))
                        {
                            if (usrWidthSet && !usrHeightSet)
                            {
                                resultHeight = inherentHeight / inherentWidth * usrWidth;
                            }
                            else if (usrHeightSet && !usrWidthSet)
                            {
                                resultWidth = inherentWidth / inherentHeight * usrHeight;
                            }
// ReSharper disable once ConditionIsAlwaysTrueOrFalse
                            else if (!usrHeightSet && !usrWidthSet)
                            {
                                resultHeight = inherentHeight;
                                resultWidth = inherentWidth;
                            }

                            if (scaleHeightSet)
                            {
                                resultHeight = resultHeight * scaleHeight;
                                resultWidth = resultWidth * scaleHeight;
                            }
                            if (scaleWidthSet)
                            {
                                resultHeight = resultHeight * scaleWidth;
                                resultWidth = resultWidth * scaleWidth;
                            }
                        }
                        else
                        {
                            if (!usrHeightSet)
                                resultHeight = inherentHeight;

                            if (!usrWidthSet)
                                resultWidth = inherentWidth;

                            if (scaleHeightSet)
                                resultHeight = resultHeight * scaleHeight;
                            if (scaleWidthSet)
                                resultWidth = resultWidth * scaleWidth;
                        }

                        formatInfo.CropWidth = (int)xPixels;
                        formatInfo.CropHeight = (int)yPixels;
                        if (_image._pictureFormat != null && !_image._pictureFormat.IsNull())
                        {
                            PictureFormat picFormat = _image.PictureFormat;
                            //Cropping in pixels.
                            XUnit cropLeft = picFormat.CropLeft.Point;
                            XUnit cropRight = picFormat.CropRight.Point;
                            XUnit cropTop = picFormat.CropTop.Point;
                            XUnit cropBottom = picFormat.CropBottom.Point;
                            formatInfo.CropX = (int)(horzRes * cropLeft.Inch);
                            formatInfo.CropY = (int)(vertRes * cropTop.Inch);
                            formatInfo.CropWidth -= (int)(horzRes * ((XUnit)(cropLeft + cropRight)).Inch);
                            formatInfo.CropHeight -= (int)(vertRes * ((XUnit)(cropTop + cropBottom)).Inch);

                            //Scaled cropping of the height and width.
                            double xScale = resultWidth / inherentWidth;
                            double yScale = resultHeight / inherentHeight;

                            cropLeft = xScale * cropLeft;
                            cropRight = xScale * cropRight;
                            cropTop = yScale * cropTop;
                            cropBottom = yScale * cropBottom;

                            resultHeight = resultHeight - cropTop - cropBottom;
                            resultWidth = resultWidth - cropLeft - cropRight;
                        }
                        if (resultHeight <= 0 || resultWidth <= 0)
                        {
                            formatInfo.Width = XUnit.FromCentimeter(2.5);
                            formatInfo.Height = XUnit.FromCentimeter(2.5);
                            Debug.WriteLine(Messages2.EmptyImageSize);
                            _failure = ImageFailure.EmptySize;
                        }
                        else
                        {
                            formatInfo.Width = resultWidth;
                            formatInfo.Height = resultHeight;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(Messages2.ImageNotReadable(_image.Name, ex.Message));
                        formatInfo.Failure = ImageFailure.NotRead;
                    }
                }
            }
            if (formatInfo.Failure != ImageFailure.None)
            {
                if (!_image._width.IsNull)
                    formatInfo.Width = _image.Width.Point;
                else
                    formatInfo.Width = XUnit.FromCentimeter(2.5);

                if (!_image._height.IsNull)
                    formatInfo.Height = _image.Height.Point;
                else
                    formatInfo.Height = XUnit.FromCentimeter(2.5);
            }
        }

        private FixedPage GetFixedPage(string xpsFilename)
        {
            var fileName = GetFileName(xpsFilename, out var index);

            return GetFixedPage(fileName, index);
        }

        private FixedPage GetFixedPage(string fileName, int index)
        {
            if (!_xpsCache.TryGetValue(fileName, out var xpsDocument))
            {
                xpsDocument = XpsDocument.Open(fileName);
                _xpsCache.Add(fileName, xpsDocument);
            }
            var fixedDocument = xpsDocument.GetDocument();
            return fixedDocument.GetFixedPage(index);
        }
    }
}
