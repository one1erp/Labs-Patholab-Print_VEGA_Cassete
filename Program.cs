//using DataMatrix.net;
//using LSEXT;
//using PrintCassete;
//using System.Drawing;
//using System.Drawing.Drawing2D;


//namespace ConsoleApp1
//{
//    class Program
//    {
//        static void Main ( string [ ] args )
//        {


//            //DatamatrixEncodingOptions encodingOptions=new DatamatrixEncodingOptions();
//            //encodingOptions.SymbolShape = ZXing.Datamatrix.Encoder.SymbolShapeHint.FORCE_SQUARE;
//            //encodingOptions.Width = 30;
//            //encodingOptions.Height = 30;

//            //encodingOptions.Margin = 1;
//            //IBarcodeWriter writer = new BarcodeWriter();
//            //writer.Format = BarcodeFormat.DATA_MATRIX;
//            //writer.Options = encodingOptions;

//            //var   result = writer.Write ( "B000001/19.1.1" );
//            //result.Save ( @"C:\Users\ashim\Desktop\NEWNEW\1.png" );
//            //var    barcodeImage = new Bitmap ( result );
//            //barcodeImage.Save ( @"C:\Users\ashim\Desktop\NEWNEW\2.png" );

//            ////   barcodeImage.Save ( @"C:\a\1.png" );

//            //return;

//            LSExtensionParameters p = null;
//            PrintCassete.PrintCasseteCls printCasseteCls=new PrintCasseteCls();
//            printCasseteCls.DEBUG = true;
//            printCasseteCls.Execute ( ref p );
//            return;


//        }

//        static void CreateBitmapAtRuntime ( )
//        {
//            return;
//            #region variables
//            var       aliqExt = ".1.1";
//            var     aliqPathoName = "P_    1-19";
//            var     aliqNautilusName = "B000001/19.1.1";
//            #endregion


//            #region תמונה ברקוד
//            DmtxImageEncoder encoder = new DmtxImageEncoder();
//            var barcodeImage = encoder.EncodeImage (aliqNautilusName/*, encoderOptions*/);
//            barcodeImage.Save ( "C:\\a\\barcodeImage.png" );
//            #endregion



//            #region תמונה טקסט
//            Bitmap stringImage = new Bitmap(80, 35);
//            Graphics flagGraphics = Graphics.FromImage(stringImage);
//            //First row
//            PointF firstLocation = new PointF(0f, 0f);
//            Font firstFont = new Font ( "Tahoma", 8,FontStyle.Bold );
//            //Second Row
//            PointF secondLocation = new PointF(40f,15f);
//            Font secondFont = new Font ( "Tahoma", 10 ,FontStyle.Regular);

//            //Draw
//            using ( Graphics graphics = Graphics.FromImage ( stringImage ) )
//            {

//                graphics.DrawString ( aliqPathoName, firstFont, Brushes.Black, firstLocation );
//                graphics.DrawString ( aliqExt, secondFont, Brushes.Black, secondLocation );

//            }
//            stringImage.Save ( "C:\\a\\stringImage.png" );

//            #endregion

//            #region תמונה משולבת

//            var mixImage=new Bitmap(108,35);
//            Graphics gmix=Graphics.FromImage(mixImage);

//            gmix.DrawImage ( barcodeImage, 0, 0, 35, 35 );
//            gmix.DrawImage ( stringImage, 38, 0f, stringImage.Width, stringImage.Height );
//            gmix.SmoothingMode = SmoothingMode.AntiAlias;
//            gmix.InterpolationMode = InterpolationMode.HighQualityBicubic;
//            gmix.PixelOffsetMode = PixelOffsetMode.HighQuality;

//            //Save the resulting image
//            mixImage.Save ( "C:\\a\\mixImage.png" );
//            #endregion

//            return;


//        }
//    }
//}
