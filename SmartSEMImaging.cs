using APILib;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using Emgu.CV.Util;
using System.Diagnostics;
//using Emgu.CV.UI;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using static System.Net.Mime.MediaTypeNames;

namespace SmartSEMImaging
{
    class SmartSEMImaging
    {
        static Dictionary<string, Image<Gray, byte>> IMGDictionary; // Dictionary for images

        static void Main(string[] args)
        {    

            Console.WriteLine("----- SmartSEM Imaging Test -----\n");

            // Initialize Api
            Api CZEMApi = new Api();

            // Define a flag to later check for initialisation
            bool apiInitialised = false;


            // Initalise communication between the CZEMApi OCX and EM Server
            //long lReturn = CZEMApi.Initialise(""); //This doesn't work on 64bit process
            long lReturn = CZEMApi.InitialiseRemoting();
            if (lReturn == 0)
            {
                apiInitialised = true;
                Console.WriteLine("Remote API correctly initialised");
            }
            else
            {
                Console.WriteLine("Remote API not initialised");
            }



            /*
           // -- TEST -- Grab image from SmartSEM API (grab) method
           if (apiInitialised)
           {
               Console.WriteLine("Testing -- Image (grab) Method\n");

               // User input section:
               Console.WriteLine("To grab EM image, please enter the following parameters.");

               Console.WriteLine("The x-offset of the image to grab (integer value. e.g.: 0): ");
               string str_ix = Console.ReadLine();
               short ix = Convert.ToInt16(str_ix);

               Console.WriteLine("The y-offset of the image to grab (integer value. e.g.: 0): ");
               string str_iy = Console.ReadLine();
               short iy = Convert.ToInt16(str_iy);

               Console.WriteLine("The width of the image to grab (integer between 0-1024): ");
               string str_il = Console.ReadLine();
               short il = Convert.ToInt16(str_il);

               Console.WriteLine("The height of the image to grab (integer between 0-768): ");
               string str_ih = Console.ReadLine();
               short ih = Convert.ToInt16(str_ih);

               Console.WriteLine("The subsampling factor for the image to grab (integer between -1-3): ");
               string str_ir = Console.ReadLine();
               short ir = Convert.ToInt16(str_ir);

               Console.WriteLine("The filename (address) where the .tiff image will be saved: ");
               Console.WriteLine("(e.g.: C:\\ProgramData\\Carl Zeiss\\SmartSEM\\Images\\Capture.tiff , but put two backslashes between folders)");
               string filename = Console.ReadLine();

               Console.WriteLine("Name the image he EM will grab: ");
               Console.WriteLine("(e.g.: Picture)");
               string value = Console.ReadLine();


               // Uncomment below values to use hard-coded values iinstead of user input v
                   short ix = 0;
                   short iy = 0;
                   short il = 1024;
                   short ih = 768;
                   short ir = -1;
                   //string filename = "C:\\Capture.tiff";
                   //string filename = "U:\\My Documents\\SmartSEM_API\\Capture.tiff";
                   string filename = "C:\\Users\\jessi\\Desktop\\SmartSEM(ECE 3970)\\Pics for ECE 3970\\Saved Grab Pics\\Capture.tiff";

                   // Set image user text
                   string value = "Picture";


               // Initialise a VariantWrapper object to reference the parameter value
               object vValue = new VariantWrapper(value);
               //CZEMApi.Set("SV_USER_TEXT", ref vValue);
               CZEMApi.Set("SV_SAMPLE_ID", ref vValue);

               // Get the stage position values
               long lStgValues = CZEMApi.Grab(ix, iy, il, ih, ir, filename);

               // Show success message, or an error message in case of an error
               switch ((ZeissErrorCode)lStgValues)
               {
                   case ZeissErrorCode.API_E_NO_ERROR:
                       Console.WriteLine("Image grabbed and saved.");
                       break;
                   case ZeissErrorCode.API_E_GRAB_FAIL:
                       Console.WriteLine("Grab command failed.");
                       break;
                   case ZeissErrorCode.API_E_NOT_INITIALISED:
                       Console.WriteLine("API not initialised.");
                       break;
               }
           } 
           */




            // -- TEST -- Get image from file, edge/shape detection, ...
            if (apiInitialised)
            {
                // User input section:
                Console.WriteLine("Enter the filename (address) where the image is: ");
                Console.WriteLine("(e.g.: C:\\ProgramData\\Carl Zeiss\\SmartSEM\\Images\\Capture.tiff , but put two backslashes between folders)");
                // **For testing use:     C:\\Users\\jessi\\Desktop\\SmartSEM (ECE 3970)\\Pics for ECE 3970\\low_res_1.tif
                // **For testing use:     C:\\Users\\jessi\\Desktop\\SmartSEM (ECE 3970)\\Pics for ECE 3970\\grab_1.tif
                string FileName = Console.ReadLine();


                // Variables for gray-scaling/edge detection/shape detection/etc
                Image<Bgr, byte> m_SourceImage    = new Image<Bgr,  byte>(FileName);               // Stores user inputted image in original color
                Image<Gray, byte> m_SourceImageGray = new Image<Gray, byte>(m_SourceImage.Size);   // Makes original source image grayscale
                Image<Gray, byte> m_InvertedImage  = new Image<Gray, byte>(m_SourceImage.Size);    // Makes source image inverted (light slides/dark backgnd)
                Image<Gray, byte> m_ThresholdImage = new Image<Gray, byte>(m_SourceImage.Size);    // Gets inverted thresholded corners
                Image<Gray, byte> m_EdgesImage     = new Image<Gray, byte>(m_SourceImage.Size);    // Gets image edges using Canny
                Image<Gray, byte> m_ContoursImage  = new Image<Gray, byte>(m_SourceImage.Size);    // Gets image contours                
                Image<Gray, byte> m_AlteredInputImage = new Image<Gray, byte>(m_SourceImage.Size); // Image to be used for image dictionary
                Image<Gray, byte> m_ShapeMatchedImage = new Image<Gray, byte>(m_SourceImage.Size); // Image that's gone through edge detection + shape contour matching

                // Convert source image to grayscale instead of Bgr
                m_SourceImageGray = m_SourceImage.Convert<Gray, byte>();

                // Sharpen image to try and better differentiate it from background
                Image<Gray, byte> m_SharpenedImage = new Image<Gray, byte>(m_SourceImage.Size);
                m_SharpenedImage = Sharpen(m_SourceImageGray, 100, 10, 10);                         

                // Inverts original user inputted m_InvertedImageimage so slides are lighter in color than the background
                //CvInvoke.Invert(m_SourceImage, m_InvertedImage, DecompMethod.LU); //LU = gaussian elimination //NOT WORKING!!
                //m_InvertedImage = m_SourceImage.Convert<Gray, Byte>().Not(); // Inverts source image after it's been converted to grayscale

                // Creates edges from image using Canny method
                CvInvoke.Canny(m_SharpenedImage, m_EdgesImage, 100, 100);

                // Applies active thresholding, which decides whether edges are present or not at an image point
                CvInvoke.AdaptiveThreshold(m_EdgesImage, m_ThresholdImage, 255, AdaptiveThresholdType.GaussianC, ThresholdType.BinaryInv, 301, 0.0);

                /* // Detect corners from image using Harris corners mthod
                Image<Gray, float> m_CornerImage = null;
                m_CornerImage = new Image<Gray, float>(m_SourceImage.Size);
                CvInvoke.CornerHarris(m_ContoursImage, m_CornerImage, 2, 3, 0.04);
                //CvInvoke.Normalize(m_CornerImage, m_CornerImage, 255, 0, Emgu.CV.CvEnum.NormType.MinMax); //Currently trash! */


                // Check all shapes and detect trapezoids
                // Copies initial image that's been altered by other image processing methods & runs trapezoid shape detection on it
                IMGDictionary = new Dictionary<string, Image<Gray, byte>>(); // initialize image dictionary
                m_AlteredInputImage = m_ThresholdImage; // copied image from canny/thresholding
                //check if image dictinary contains input key already 
                if (IMGDictionary.ContainsKey("input"))
                {
                    IMGDictionary.Remove("input");
                }
                IMGDictionary.Add("input", m_AlteredInputImage);

                // apply shape matching method
                m_ShapeMatchedImage = ApplyShapeMatching();


                //**TODO -- after detecting all trapezoids, get corners and store corners with respective trapezoid


                // Displays altered image
                Image<Gray, byte> m_OutputImage = m_ShapeMatchedImage;
                String win1 = "Test Window"; //The name of the window
                CvInvoke.NamedWindow(win1); //Create the window using the specific name
                CvInvoke.Imshow(win1, m_OutputImage); //Show the image
                //CvInvoke.Resize(m_ContoursImage, m_OutputImage, new Size(1024, 768));
                CvInvoke.WaitKey(0);  //Wait for the key pressing event
                CvInvoke.DestroyWindow(win1); //Destroy the window if key is pressed
               
                
            }//End of If(apiInitialized)


        }//End of Main


        /// <summary>
        /// Mathod that sharpens an input image.
        /// Algorithm to sharpen an image:
        /// 1. Blur the original image using the Gaussian filter with given Mask Size & Sigma
        /// 2. Subtract the blurred image from the original (result is called Mask) to eliminate background and get the edges regions
        /// 3. Add a weighted portion from the mask to the original image by multiplying the Mask (the edges only) by K to enhance edges regions
        /// </summary>
        /// <param name="image"> Input image </param>
        /// <param name="sigma1"></param>
        /// <param name="sigma2"></param>
        /// <param name="k"> User input (If K = 1 Unsharp, If K > 1 Highboost) </param>
        /// <returns> Sharpened image </returns>
        private static Image<Gray, byte> Sharpen (Image<Gray, byte> image, double sigma1, double sigma2, int k)
        {
            var h = image.Height;
            var w = image.Width;
            h = (h % 2 == 0) ? h - 1 : h;
            w = (w % 2 == 0) ? w - 1 : w;

            // apply gaussian smoothing using w, h and sigma 
            var gaussianSmooth = image.SmoothGaussian(w, h, sigma1, sigma2);
            // obtain the mask by subtracting the gaussian smoothed image from the original one 
            // Mask(x,y) = Orig(x,y) – Blurred(x,y)
            var mask = image - gaussianSmooth;
            // add a weighted value k to the obtained mask 
            mask *= k;
            // sum with the original image 
            image += mask;
            // Result(x,y) = Orig(x,y) + K × Mask(x,y)
            return image;
        }


        /// <summary>
        /// Method to match input image shapes with a template to locate trapezoids in the image.
        /// </summary>
        /// <param name="shape_threshold"></param>
        /// <exception cref="Exception"></exception>
        private static Image<Gray, byte> ApplyShapeMatching(double shape_threshold = 0.1)
        {
            string shapefile = "C:\\Users\\jessi\\Desktop\\SmartSEM (ECE 3970)\\Pics for ECE 3970\\trapezoid_reference.tif"; // address for template image (trapezoid shape)
            Image<Gray, byte> m_ShapeTemplate = new Image<Gray, byte>(shapefile); //image that's a template to compare against image in image dictionary

            try
            {
                if (IMGDictionary["input"] == null)
                {
                    throw new Exception("Error: Select an input image");
                }

                Image<Gray, byte> m_DictClone = IMGDictionary["input"].Clone(); //copies the input image in image dictionary
                var m_DictContours = CalculateContours(m_DictClone); //gets the contours for input image in image dictionary
                var m_ShapeContours = CalculateContours(m_ShapeTemplate); //gets the contours for shape template image

                if (m_DictContours.Size == 0 || m_ShapeContours.Size == 0)
                {
                    throw new Exception("Not enough contours");
                }

                for (int i = 0; i < m_DictContours.Size; i++)
                {
                    var distance = CvInvoke.MatchShapes(m_DictContours[i], m_ShapeContours[0], ContoursMatchType.I2);

                    if (distance <= shape_threshold)
                    {
                        var rect = CvInvoke.BoundingRectangle(m_DictContours[i]);
                        m_DictClone.Draw(rect, new Gray(), 4);
                        CvInvoke.PutText(m_DictClone, distance.ToString("F6"), new Point(rect.X, rect.Y+20), FontFace.HersheyPlain, 3, new MCvScalar(255, 0, 0));
                    }
                }

                return m_DictClone;

            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }


        /// <summary>
        /// Method that calculates the contours of an input image.
        /// </summary>
        /// <param name="image"> Input image </param>
        /// <param name="thresholdArea"> Hard-coded threshold amount </param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static VectorOfVectorOfPoint CalculateContours (Image<Gray, byte> image, double thresholdArea = 10000)
        {
            try
            {
                VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();
                VectorOfVectorOfPoint filteredContours = new VectorOfVectorOfPoint();
                Mat hierarchy = new Mat();

                CvInvoke.FindContours(image, contours, hierarchy, RetrType.External, ChainApproxMethod.ChainApproxSimple);                             

                for (int i = 0; i < contours.Size; i++)
                {
                    var area = CvInvoke.ContourArea(contours[i]);
                    if (area >= thresholdArea)
                    {
                        filteredContours.Push(contours[i]);
                    }
                }

                return filteredContours;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
        

        // Some SmartSEM API error codes
        public enum ZeissErrorCode
        {
            API_E_NO_ERROR = 0,

            // Failed to translate parameter into an id
            API_E_GET_TRANSLATE_FAIL = 1000,

            // Failed to translate command into an id
            API_E_EXEC_TRANSLATE_FAIL = 1011,

            // Failed to execute command
            API_E_EXEC_CMD_FAIL = 1012,

            // Failed to execute file macro
            API_E_EXEC_MCF_FAIL = 1013,

            // Failed to execute library macro
            API_E_EXEC_MCL_FAIL = 1014,

            // Command supplied is not implemented
            API_E_EXEC_BAD_COMMAND = 1015,

            // Grab command failed
            API_E_GRAB_FAIL = 1016,

            // API not initialised
            API_E_NOT_INITIALISED = 1019,

            // Get limits failed
            API_E_GET_LIMITS_FAIL = 1022,
        }

    }
}

