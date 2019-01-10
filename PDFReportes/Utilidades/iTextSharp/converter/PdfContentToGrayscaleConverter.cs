#region This file is part of the iTextSharp.extended library
//
// PdfContentToGrayscaleConverter.cs
//
// Author: Josip Habjan (habjan@gmail.com, http://www.linkedin.com/in/habjan) 
// Copyright (c) 2013 Josip Habjan. All rights reserved.
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>. 
//
#endregion

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf.modifier;

namespace iTextSharp.text.pdf.converter
{
	public class PdfContentToGrayscaleConverter
    {
        #region Private variables

        private static ColorMatrix _colorMatrix = new ColorMatrix(
                            new float[][] 
                            {
                                new float[] {.3f, .3f, .3f, 0, 0},
                                new float[] {.59f, .59f, .59f, 0, 0},
                                new float[] {.11f, .11f, .11f, 0, 0},
                                new float[] {0, 0, 0, 1, 0},
                                new float[] {0, 0, 0, 0, 1}
                            });

        private PdfContentModifier _modifier = new PdfContentModifier();
        private List<int> _convertedIndirectList = new List<int>();

        #endregion

        #region Constructor

        public PdfContentToGrayscaleConverter()
        {
            _modifier.RegisterOperator("G", SetStrokingGray);
            _modifier.RegisterOperator("g", SetNonStrokingGray);
            _modifier.RegisterOperator("RG", SetStrokingRGB);
            _modifier.RegisterOperator("rg", SetNonStrokingRGB);
            _modifier.RegisterOperator("K", SetStrokingCMYK);
            _modifier.RegisterOperator("k", SetNonStrokingCMYK);
            _modifier.RegisterOperator("CS", SetStrokingGeneral);
            _modifier.RegisterOperator("cs", SetNonStrokingGeneral);
            _modifier.RegisterOperator("SC", SetStrokingGeneral);
            _modifier.RegisterOperator("sc", SetNonStrokingGeneral);
            _modifier.RegisterOperator("SCN", SetStrokingGeneral);
            _modifier.RegisterOperator("scn", SetNonStrokingGeneral);

            _modifier.RegisterOperator("Do", Do);
        }

        #endregion

        #region Convert

        public void Convert(PdfReader reader, int pageNumber)
        {
            _modifier.Modify(reader, pageNumber);
        }

        #endregion

        #region SetStrokingGray

        private List<PdfObject> SetStrokingGray(PdfLiteral oper, List<PdfObject> operands)
        {
            // it's already gray, nothing to do here, return what we got
            return operands;
        }

        #endregion

        #region SetNonStrokingGray

        private List<PdfObject> SetNonStrokingGray(PdfLiteral oper, List<PdfObject> operands)
        {
            // it's already gray, nothing to do here, return what we got
            return operands;
        }

        #endregion

        #region SetStrokingRGB

        private List<PdfObject> SetStrokingRGB(PdfLiteral oper, List<PdfObject> operands)
        {
            // convert RGB to Grayscale
            GrayColor color = this.Convert_RGB_To_Grayscale(
                ((PdfNumber)operands[0]).FloatValue,
                ((PdfNumber)operands[1]).FloatValue,
                ((PdfNumber)operands[2]).FloatValue);

            return new List<PdfObject>() 
                { 
                    new PdfNumber(color.Gray),
                    new PdfLiteral("G")
                };
        }

        #endregion

        #region SetNonStrokingRGB

        private List<PdfObject> SetNonStrokingRGB(PdfLiteral oper, List<PdfObject> operands)
        {
            // convert RGB to Grayscale
            GrayColor color = this.Convert_RGB_To_Grayscale(
                ((PdfNumber)operands[0]).FloatValue,
                ((PdfNumber)operands[1]).FloatValue,
                ((PdfNumber)operands[2]).FloatValue);

            return new List<PdfObject>() 
                { 
                    new PdfNumber(color.Gray),
                    new PdfLiteral("g")
                };
        }

        #endregion

        #region SetStrokingCMYK

        private List<PdfObject> SetStrokingCMYK(PdfLiteral oper, List<PdfObject> operands)
        {
            // convert CMYK to Grayscale
            GrayColor color = this.Convert_CMYK_To_Grayscale(
                ((PdfNumber)operands[0]).FloatValue,
                ((PdfNumber)operands[1]).FloatValue,
                ((PdfNumber)operands[2]).FloatValue,
                ((PdfNumber)operands[3]).FloatValue);

            return new List<PdfObject>() 
                { 
                    new PdfNumber(color.Gray),
                    new PdfLiteral("G")
                };
        }

        #endregion

        #region SetNonStrokingCMYK

        private List<PdfObject> SetNonStrokingCMYK(PdfLiteral oper, List<PdfObject> operands)
        {
            // convert CMYK to Grayscale
            GrayColor color = this.Convert_CMYK_To_Grayscale(
                ((PdfNumber)operands[0]).FloatValue,
                ((PdfNumber)operands[1]).FloatValue,
                ((PdfNumber)operands[2]).FloatValue,
                ((PdfNumber)operands[3]).FloatValue);

            return new List<PdfObject>() 
                { 
                    new PdfNumber(color.Gray),
                    new PdfLiteral("g")
                };
        }

        #endregion

        #region SetStrokingGeneral

        private List<PdfObject> SetStrokingGeneral(PdfLiteral oper, List<PdfObject> operands)
        {
            if (operands.Count == 4)
            {
                PdfNumber r = (PdfNumber)operands[0];
                PdfNumber g = (PdfNumber)operands[1];
                PdfNumber b = (PdfNumber)operands[2];

                BaseColor color = new BaseColor(r.FloatValue, g.FloatValue, b.FloatValue);
                GrayColor grayColor = this.Convert_RGB_To_Grayscale(color);

                return new List<PdfObject>() 
                { 
                    new PdfNumber(grayColor.Gray),
                    new PdfLiteral("G")
                };
            }
            else
            {
                return new List<PdfObject>() 
                { 
                    new PdfNumber(GrayColor.GRAYBLACK.Gray),
                    new PdfLiteral("G")
                };
            }
        }

        #endregion

        #region SetNonStrokingGeneral

        private List<PdfObject> SetNonStrokingGeneral(PdfLiteral oper, List<PdfObject> operands)
        {
            if (operands.Count == 4)
            {
                PdfNumber r = (PdfNumber)operands[0];
                PdfNumber g = (PdfNumber)operands[1];
                PdfNumber b = (PdfNumber)operands[2];

                BaseColor color = new BaseColor(r.FloatValue, g.FloatValue, b.FloatValue);
                GrayColor grayColor = this.Convert_RGB_To_Grayscale(color);

                return new List<PdfObject>() 
                { 
                    new PdfNumber(grayColor.Gray),
                    new PdfLiteral("g")
                };
            }
            else
            {
                //PdfDictionary o = _modifier.ResourceDictionary.GetAsDict(PdfName.COLORSPACE);
                //PdfObject po = o.GetDirectObject(operands[0] as PdfName);

                return new List<PdfObject>() 
                { 
                    new PdfNumber(GrayColor.GRAYBLACK.Gray),
                    new PdfLiteral("g")
                };
            }
        }

        #endregion

        #region Do

        private List<PdfObject> Do(PdfLiteral oper, List<PdfObject> operands)
        {
            PdfName xobjectName = (PdfName)operands[0];

            System.Diagnostics.Debug.WriteLine("O: " + xobjectName.ToString());

            PdfDictionary xobjects = _modifier.ResourceDictionary.GetAsDict(PdfName.XOBJECT);
            PdfObject po = xobjects.Get(xobjectName);

            int n = (po as PdfIndirectReference).Number;

            if (_convertedIndirectList.Contains(n))
            {
                return operands;
            }
            else
            {
                _convertedIndirectList.Add(n);
            }

            PdfObject xobject = xobjects.GetDirectObject(xobjectName);

            if (xobject.IsStream())
            {
                PdfStream xobjectStream = (PdfStream)xobject;
                PdfName subType = xobjectStream.GetAsName(PdfName.SUBTYPE);

                if (subType == PdfName.FORM)
                {
                    this.Do_Form(xobjectStream);
                }
                else if (subType == PdfName.IMAGE)
                {
                    this.Do_Image(xobjectStream);
                }
            }

            return operands;
        }

        #endregion

        #region Do_Form

        private void Do_Form(PdfStream stream)
        {
            PdfDictionary resources = stream.GetAsDict(PdfName.RESOURCES);

            byte[] contentBytes = ContentByteUtils.GetContentBytesFromContentObject(stream);

            contentBytes = _modifier.Modify(contentBytes, resources);

            PRStream prStream = stream as PRStream;

            prStream.SetData(contentBytes);
        }

        #endregion

        #region Do_Image

        private void Do_Image(PdfStream stream)
        {
            byte[] imageBuffer = null;

            PdfName filter = PdfName.NONE;

            if (stream.Contains(PdfName.FILTER))
            {
                filter = stream.GetAsName(PdfName.FILTER);
            }

            int imageWidth = stream.GetAsNumber(PdfName.WIDTH).IntValue;
            int imageHeight = stream.GetAsNumber(PdfName.HEIGHT).IntValue;
            int imageBpp = stream.GetAsNumber(PdfName.BITSPERCOMPONENT).IntValue;

            PRStream prStream = stream as PRStream;

            bool cannotReadImage = false;

            Bitmap image = null;

            try
            {
                PdfImageObject pdfImage = new PdfImageObject(prStream);
                image = pdfImage.GetDrawingImage() as Bitmap;
            }
            catch
            {
                try
                {
                    if (filter == PdfName.FLATEDECODE)
                    {
                        byte[] streamBuffer = PdfReader.GetStreamBytes(prStream);
                        image = this.CreateBitmapFromFlateDecodeImage(streamBuffer, imageWidth, imageHeight, imageBpp);
                    }
                }
                catch
                {
                    cannotReadImage = true;
                }
            }

            if (!cannotReadImage)
            {
                image = this.ConvertToGrayscale(image);

                using (var ms = new MemoryStream())
                {
                    System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Quality;

                    ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Jpeg);                    
                    EncoderParameters myEncoderParameters = new EncoderParameters(1);

                    //EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 80L);                    
                    EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 15L);
                    myEncoderParameters.Param[0] = myEncoderParameter;

                    image.Save(ms, jgpEncoder, myEncoderParameters);

                    imageBuffer = ms.GetBuffer();
                }

                Image compressedImage = Image.GetInstance(imageBuffer);

                imageBuffer = compressedImage.OriginalData;

                prStream.Clear();

                //prStream.SetData(imageBuffer, false, PRStream.NO_COMPRESSION);                
                prStream.SetData(imageBuffer, false, PRStream.BEST_COMPRESSION);
                prStream.Put(PdfName.TYPE, PdfName.XOBJECT);
                prStream.Put(PdfName.SUBTYPE, PdfName.IMAGE);
                prStream.Put(PdfName.FILTER, PdfName.DCTDECODE);
                prStream.Put(PdfName.WIDTH, new PdfNumber(image.Width));
                prStream.Put(PdfName.HEIGHT, new PdfNumber(image.Height));
                prStream.Put(PdfName.BITSPERCOMPONENT, new PdfNumber(8));
                prStream.Put(PdfName.COLORSPACE, PdfName.DEVICERGB);                
                prStream.Put(PdfName.LENGTH, new PdfNumber(imageBuffer.LongLength));
            }
        }

        #endregion

        #region CreateBitmapFromFlateDecodeImage

        private Bitmap CreateBitmapFromFlateDecodeImage(byte[] buffer, int width, int height, int bpp)
        {
            PixelFormat pixelFormat;

            switch (bpp)
            {
                case 1:
                    {
                        pixelFormat = PixelFormat.Format1bppIndexed;
                        break;
                    }
                case 8:
                    {
                        pixelFormat = PixelFormat.Format8bppIndexed;
                        break;
                    }
                default:
                    {
                        pixelFormat = PixelFormat.Format24bppRgb;
                        break;
                    }
            }

            using (Bitmap bmp = new Bitmap(width, height, pixelFormat))
            {
                var bmpData = bmp.LockBits(new System.Drawing.Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, pixelFormat);

                int length = (int)Math.Ceiling(width * bpp / 8.0);

                for (int i = 0; i < height; i++)
                {
                    int offset = i * length;
                    int scanOffset = i * bmpData.Stride;

                    Marshal.Copy(buffer, offset, new IntPtr(bmpData.Scan0.ToInt32() + scanOffset), length);
                }

                bmp.UnlockBits(bmpData);

                return bmp.Clone() as Bitmap;
            }
        }

        #endregion

        #region GetEncoder

        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();

            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }

            return null;
        }

        #endregion

        #region ConvertToGrayscale

        private Bitmap ConvertToGrayscale(Bitmap original)
        {
            Bitmap newBitmap = new Bitmap(original.Width, original.Height);
            Graphics g = Graphics.FromImage(newBitmap);
            ImageAttributes attributes = new ImageAttributes();
            attributes.SetColorMatrix(_colorMatrix);
            g.DrawImage(original, new System.Drawing.Rectangle(0, 0, original.Width, original.Height), 0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);
            g.Dispose();
            return newBitmap;
        }

        #endregion

        #region CMYK_To_RGB

        private BaseColor Convert_CMYK_To_RGB(float c, float m, float y, float k)
        {
            int red = (int)((1 - c) * (1 - k) * 255.0);
            int green = (int)((1 - m) * (1 - k) * 255.0);
            int blue = (int)((1 - y) * (1 - k) * 255.0);

            return new BaseColor(red, green, blue);
        }

        #endregion

        #region Convert_CMYK_To_Grayscale

        private GrayColor Convert_CMYK_To_Grayscale(float c, float m, float y, float k)
        {
            BaseColor color = this.Convert_CMYK_To_RGB(c, m, y, k);
            return this.Convert_RGB_To_Grayscale(color);
        }

        #endregion

        #region Convert_RGB_To_Grayscale

        private GrayColor Convert_RGB_To_Grayscale(float r, float g, float b)
        {
            return this.Convert_RGB_To_Grayscale(new BaseColor(r, g, b));
        }

        #endregion

        #region Convert_RGB_To_Grayscale

        private GrayColor Convert_RGB_To_Grayscale(BaseColor color)
        {
            return new GrayColor((int)((color.R * 0.30f) + (color.G * 0.59) + (color.B * 0.11)));
        }
                
        #endregion
    }
}
