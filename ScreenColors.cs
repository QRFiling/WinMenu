using Litmus;
using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinMenu
{
    class ScreenColors
    {
        static ILitmus litmus = new DirectxColorProvider();

        public static async Task<Tuple<System.Windows.Media.Color, System.Windows.Media.Color, System.Windows.Media.SolidColorBrush>> GetColors()
        {
            LitmusColor c = null;
            await Task.Run(() => { c = litmus.GetAverageColor(); });

            Color color = UpdateColor(Color.FromArgb(c.Red, c.Green, c.Blue));
            Color color2 = ControlPaint.Light(color);
            Color color3 = ControlPaint.LightLight(color);

            return new Tuple<System.Windows.Media.Color, System.Windows.Media.Color,
                System.Windows.Media.SolidColorBrush>(ConvertColor(color), ConvertColor(color2),
                new System.Windows.Media.SolidColorBrush(ConvertColor(color3)));
        }

        static Color UpdateColor(Color color)
        {
            if (color.GetBrightness() < 0.7)
                color = ControlPaint.LightLight(ControlPaint.LightLight(color));

            color = ControlPaint.Light(color);
            color = Blend(Color.FromArgb(200, 235, 250), color, 0.2);

            return color;
        }

        static Color Blend(Color color, Color backColor, double amount)
        {
            byte r = (byte)(color.R * amount + backColor.R * (1 - amount));
            byte g = (byte)(color.G * amount + backColor.G * (1 - amount));
            byte b = (byte)(color.B * amount + backColor.B * (1 - amount));

            return Color.FromArgb(r, g, b);
        }

        static System.Windows.Media.Color ConvertColor(Color color) =>
            System.Windows.Media.Color.FromRgb(color.R, color.G, color.B);
    }
}
