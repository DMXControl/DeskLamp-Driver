﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Drawing;

namespace DeskLamp
{
    class DeskLampTest
    {
        static void Main(string[] args)
        {
            IEnumerable<string> lamps = DeskLamp.DeskLampInstance.GetAvailableDeskLamps();

            //No Arguments, run Test
            if(args == null || args.Length == 0)
            {
                PrintDesklampList(lamps);
                RunDesklampTest(lamps);
            }
            else //Write RGB / Adapter Property into Desklamps
            {
                int indexSetRgb = IndexOf(args, c => c.Equals("-setrgb", StringComparison.InvariantCultureIgnoreCase));
                int indexSetAdapter = IndexOf(args, c => c.Equals("-setadapter", StringComparison.InvariantCultureIgnoreCase));
                if (indexSetRgb == -1 && indexSetAdapter == -1) // not found
                {
                    PrintParameterHelp();
                    return;
                }
                bool? rgbPara = null;
                bool? adapterPara = null;
                if (indexSetRgb != -1)
                {
                    bool v;
                    string tmp = args.ElementAtOrDefault(indexSetRgb + 1);
                    if (String.IsNullOrEmpty(tmp) || !Boolean.TryParse(tmp, out v))
                    {
                        PrintParameterHelp();
                        return;
                    }
                    rgbPara = v;
                }
                if (indexSetAdapter != -1)
                {
                    bool v;
                    string tmp = args.ElementAtOrDefault(indexSetAdapter + 1);
                    if (String.IsNullOrEmpty(tmp) || !Boolean.TryParse(tmp, out v))
                    {
                        PrintParameterHelp();
                        return;
                    }
                    adapterPara = v;
                }
                int indexSetID = IndexOf(args, c => c.Equals("-id", StringComparison.InvariantCultureIgnoreCase));
                string paraid = null;
                if (indexSetID != -1) //something found
                {
                    paraid = args.ElementAtOrDefault(indexSetID + 1);
                    if (String.IsNullOrEmpty(paraid))
                    {
                        PrintParameterHelp();
                        return;
                    }
                }

                PrintDesklampList(lamps);

                if (String.IsNullOrEmpty(paraid))
                {
                    if (rgbPara != null)
                        Console.WriteLine("No Desklamp ID provided with \"-id\" parameter so all connected Desklamps RGB Mode will be set to \"{0}\"", rgbPara);
                    if (adapterPara != null)
                        Console.WriteLine("No Desklamp ID provided with \"-id\" parameter so all connected Desklamps Adapter Mode will be set to \"{0}\"", adapterPara);
                }
                else if (lamps.Contains(paraid))
                {
                    if (rgbPara != null)
                        Console.WriteLine("Desklamp ID \"{0}\" provided with \"-id\" parameter RGB Mode will be set to \"{1}\"", paraid, rgbPara);
                    if (adapterPara != null)
                        Console.WriteLine("Desklamp ID \"{0}\" provided with \"-id\" parameter Adapter Mode will be set to \"{1}\"", paraid, adapterPara);
                    lamps = new string[] { paraid };
                }
                else
                {
                    Console.WriteLine("Desklamp ID \"{0}\" provided with \"-id\" parameter not connected!", paraid);
                    lamps = Enumerable.Empty<string>();
                }

                int programmed = 0, errors = 0;
                foreach (string id in lamps)
                {
                    System.Console.WriteLine();
                    DeskLamp.DeskLampInstance lamp = new DeskLampInstance(id);
                    if (lamp.IsAvailable)
                    {
                        bool? ergRgb = null, ergAdapter = null;
                        if (rgbPara != null)
                            ergRgb = SetAndValidateRGB(lamp, (bool)rgbPara);
                        if (adapterPara != null)
                            ergAdapter = SetAndValidateAdapter(lamp, (bool)adapterPara);

                        if (ergRgb == true || ergAdapter == true)
                            programmed++;
                        if (ergRgb == false || ergAdapter == false)
                            errors++;
                    }
                    else
                    {
                        System.Console.WriteLine("Lamp with ID \"{0}\" not available!", lamp.ID);
                    }
                }
                Console.WriteLine("Setting RGB / Adapter Mode done! {0} Lamps programmed. {1} Errors.", programmed, errors);
                Console.WriteLine("Programm will exit in 5 seconds.");
                System.Threading.Thread.Sleep(5000);
            }
        }

        static bool? SetAndValidateRGB(DeskLamp.DeskLampInstance lamp, bool rgbValue)
        {
            if (!lamp.HasRGB)
            {
                Console.WriteLine("Lamp \"{0}\" doesn't support RGB", lamp.ID);
            }
            else if (lamp.IsRGB != rgbValue)
            {
                lamp.IsRGB = rgbValue;
                System.Threading.Thread.Sleep(100); //AL 2016-06-30: Don't know whether this Sleep is required...
                if (lamp.IsRGB == rgbValue)
                {
                    Console.WriteLine("Successfull set RGB Mode of Lamp \"{0}\" to \"{1}\"", lamp.ID, rgbValue);
                    return true;
                }
                else
                {
                    Console.WriteLine("Error setting RGB Mode of Lamp \"{0}\" to \"{1}\"", lamp.ID, rgbValue);
                    return false;
                }
            }
            else
                Console.WriteLine("Lamp \"{0}\" RGB Mode is already \"{1}\"", lamp.ID, rgbValue);
            return null;
        }

        static bool? SetAndValidateAdapter(DeskLamp.DeskLampInstance lamp, bool adapterValue)
        {
            if (lamp.Version < 2)
            {
                Console.WriteLine("Lamp \"{0}\" is HW Rev. 1 and is always an Adapter", lamp.ID);
            }
            else if (lamp.IsAdapter != adapterValue)
            {
                lamp.IsAdapter = adapterValue;
                System.Threading.Thread.Sleep(100); //AL 2016-06-30: Don't know whether this Sleep is required...
                if (lamp.IsAdapter == adapterValue)
                {
                    Console.WriteLine("Successfull set Adapter Mode of Lamp \"{0}\" to \"{1}\"", lamp.ID, adapterValue);
                    return true;
                }
                else
                {
                    Console.WriteLine("Error setting Adapter Mode of Lamp \"{0}\" to \"{1}\"", lamp.ID, adapterValue);
                    return false;
                }
            }
            else
                Console.WriteLine("Lamp \"{0}\" Adapter Mode is already \"{1}\"", lamp.ID, adapterValue);
            return null;
        }

        static void PrintParameterHelp()
        {
            Console.WriteLine("DeskLampTest.exe");
            Console.WriteLine("-setrgb [true | false]: Set RGB Mode of connected Desklamp to true / false");
            Console.WriteLine("-setadapter [true | false]: Set Adapter Mode of connected Desklamp to true / false");
            Console.WriteLine("-id [DesklampID]: Optional ID for Set RGB Mode");
        }

        static void PrintDesklampList(IEnumerable<string> lamps)
        {
            if (lamps == null)
                return;

            int cnt = lamps.Count();
            Console.WriteLine();
            Console.WriteLine("Detected {0} DeskLamps.", cnt);
            if (cnt > 0)
            {
                Console.WriteLine("Listing IDs:");
                foreach (string id in lamps)
                {
                    Console.WriteLine(" * {0}", id);
                }
            }
            Console.WriteLine();
        }

        static void RunDesklampTest(IEnumerable<string> lamps)
        {
            int tested = 0, testedrgb = 0, notavailable = 0;
            foreach (string id in lamps ?? Enumerable.Empty<string>())
            {
                System.Console.WriteLine();
                DeskLamp.DeskLampInstance lamp = new DeskLampInstance(id);
                if (lamp.IsAvailable)
                {
                    System.Console.WriteLine("Lamp with ID \"{0}\" is available", lamp.ID);
                    System.Console.WriteLine("Lamp version: {0}", lamp.Version);
                    System.Console.WriteLine("Lamp is Adapter: {0}", lamp.IsAdapter);

                    if (!lamp.ExternalUSBConnected)
                    {
                        if (lamp.HasExternalUSBDetection)
                            System.Console.WriteLine("No intelligent USB device detected, dimming ok");
                        else
                            System.Console.WriteLine("Warn: Desklamp not capable of detecting external USB device!");

                        System.Console.Write("Fading brightness...");
                        int dir = 1;
                        lamp.Brightness = 0;
                        for (int i = 0; i < 6; ++i)
                        {
                            for (int j = 0; j < 255; ++j)
                            {
                                lamp.Brightness += (byte)dir;
                                System.Threading.Thread.Sleep(1);
                            }
                            dir = -dir;
                        }
                        System.Console.WriteLine(" Done.");

                        if (lamp.HasStrobe)
                        {
                            System.Console.WriteLine("Setting strobe speed");
                            lamp.Brightness = 255;
                            lamp.Strobe = 192;
                            System.Threading.Thread.Sleep(3000);
                            System.Console.WriteLine("Disabling strobe");
                            lamp.Strobe = 0;
                        }

                        if (lamp.IsRGB)
                        {
                            System.Console.WriteLine("Lamp is RGB capable");
                            System.Console.WriteLine("Current color: {0}", lamp.Color);
                            System.Console.Write("Cycling through rainbow...");
                            for (double i = 0; i < 1; i += 0.01)
                            {
                                Color c = HSL2RGB(i, 0.5, 0.5);
                                lamp.Color = c;
                                System.Threading.Thread.Sleep(100);
                            }
                            System.Console.WriteLine(" Done.");
                            lamp.Color = Color.White;
                            testedrgb++;
                        }
                        else
                        {
                            System.Console.WriteLine("Lamp is single-channel");
                        }
                        tested++;
                    }
                    else
                    {
                        System.Console.WriteLine("External intelligent USB device detected, dimming disabled!");
                    }
                }
                else
                {
                    System.Console.WriteLine("Lamp with ID \"{0}\" not available!", lamp.ID);
                    notavailable++;
                }
            }
            System.Console.WriteLine("Test done! {0} Lamps tested ({1} with RGB). {2} not available.", tested, testedrgb, notavailable);
            Console.WriteLine("Programm will exit in 5 seconds.");
            System.Threading.Thread.Sleep(5000);
        }

        // Given H,S,L in range of 0-1
        // Returns a Color (RGB struct) in range of 0-255
        static Color HSL2RGB(double h, double sl, double l) {
            double v;
            double r, g, b;

            r = l;   // default to gray
            g = l;
            b = l;
            v = (l <= 0.5) ? (l * (1.0 + sl)) : (l + sl - l * sl);
            if (v > 0) {
                double m;
                double sv;
                int sextant;
                double fract, vsf, mid1, mid2;

                m = l + l - v;
                sv = (v - m) / v;
                h *= 6.0;
                sextant = (int)h;
                fract = h - sextant;
                vsf = v * sv * fract;
                mid1 = m + vsf;
                mid2 = v - vsf;
                switch (sextant) {
                    case 0:
                        r = v;
                        g = mid1;
                        b = m;
                        break;
                    case 1:
                        r = mid2;
                        g = v;
                        b = m;
                        break;
                    case 2:
                        r = m;
                        g = v;
                        b = mid1;
                        break;
                    case 3:
                        r = m;
                        g = mid2;
                        b = v;
                        break;
                    case 4:
                        r = mid1;
                        g = m;
                        b = v;
                        break;
                    case 5:
                        r = v;
                        g = m;
                        b = mid2;
                        break;
                }
            }
            return Color.FromArgb(Convert.ToInt32(r * 255.0f), Convert.ToInt32(g * 255.0f), Convert.ToInt32(b * 255.0f));
        }

        static int IndexOf<T>(IEnumerable<T> source, Predicate<T> predicate)
        {
            if (predicate == null)
                throw new ArgumentNullException("predicate");
            int i = 0;
            foreach (T t in source ?? Enumerable.Empty<T>())
            {
                if (predicate(t))
                    return i;
                i++;
            }
            return -1;
        }
    }
}
