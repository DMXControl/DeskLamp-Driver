using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace DeskLamp
{
    class DeskLampTest
    {
        static void Main(string[] args)
        {
            List<string> lamps = DeskLamp.DeskLampInstance.GetAvailableDeskLamps();
            System.Console.WriteLine("Detected {0} DeskLamps. Listing IDs:", lamps.Count);
            foreach(string id in lamps) {
                System.Console.WriteLine(" * {0}", id);
            }

            foreach (string id in lamps) {
                System.Console.WriteLine();
                DeskLamp.DeskLampInstance lamp = new DeskLampInstance(id);
                if (lamp.IsAvailable) {
                    System.Console.WriteLine("Lamp with ID {0} is available", lamp.ID);
                    System.Console.WriteLine("Lamp version: {0}", lamp.Version);

                    if (!lamp.ExternalUSBConnected) {
                        System.Console.WriteLine("No intelligent USB device detected, dimming ok");

                        System.Console.Write("Fading brightness...");
                        int dir = 1;
                        lamp.Brightness = 0;
                        for (int i = 0; i < 6; ++i) {
                            for (int j = 0; j < 255; ++j) {
                                lamp.Brightness += (byte)dir;
                                System.Threading.Thread.Sleep(1);
                            }
                            dir = -dir;
                        }
                        System.Console.WriteLine(" Done.");

                        if (lamp.Version >= 2) {
                            System.Console.WriteLine("Setting strobe speed");
                            lamp.Brightness = 255;
                            lamp.Strobe = 192;
                            System.Threading.Thread.Sleep(3000);
                            System.Console.WriteLine("Disabling strobe");
                            lamp.Strobe = 0;
                        }

                        if (lamp.IsRGB) {
                            System.Console.WriteLine("Lamp is RGB capable");
                            System.Console.WriteLine("Current color: {0}", lamp.Color);
                            System.Console.Write("Cycling through rainbow...");
                            for (double i = 0; i < 1; i += 0.01) {
                                Color c = HSL2RGB(i, 0.5, 0.5);
                                lamp.Color = c;
                                System.Threading.Thread.Sleep(100);
                            }
                            System.Console.WriteLine(" Done.");
                            lamp.Color = Color.White;
                        } else {
                            System.Console.WriteLine("Lamp is single-channel");
                        }
                    } else {
                        System.Console.WriteLine("External intelligent USB device detected, dimming disabled!");
                    }
                } else {
                    System.Console.WriteLine("Lamp with ID {0} not available!", lamp.ID);
                }
            }
        }

        // Given H,S,L in range of 0-1
        // Returns a Color (RGB struct) in range of 0-255
        public static Color HSL2RGB(double h, double sl, double l) {
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
    }
}
