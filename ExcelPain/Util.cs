using System;
using System.Drawing;

namespace ExcelPain
{
    public static class Util
    {
        /// <summary>
        /// resize a bitmap. the original bitmap is disposed
        /// </summary>
        /// <param name="org">the original bitmap</param>
        /// <param name="w">new width</param>
        /// <param name="h">new height</param>
        /// <returns>scaled bitmap</returns>
        public static Bitmap Resize(this Bitmap org, int w, int h)
        {
            Bitmap scaled = new Bitmap(w, h);

            // get graphics from the scaled bitmap
            // and draw bitmap scaled
            using (Graphics g = Graphics.FromImage(scaled))
                g.DrawImage(org, 0, 0, w, h);

            org.Dispose();
            return scaled;
        }

        /// <summary>
        /// Get excel collumn name for a number. 
        /// 1 -> A
        /// 2 -> B
        /// ...
        /// </summary>
        /// <param name="columnNumber">the collumn number, >1</param>
        /// <returns>collumn name</returns>
        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// get the end location (bottom right) of a rectangle
        /// </summary>
        /// <param name="r">the rectangle</param>
        /// <returns>end location</returns>
        public static Point EndLocation(this Rectangle r)
        {
            return new Point(r.X + r.Width, r.Y + r.Height);
        }

        /// <summary>
        /// convert a point to a excel cell.
        /// coords are converted from 0- index to 1- index (X +1 and Y +1)
        /// </summary>
        /// <param name="p">the point</param>
        /// <returns>the excel cell name</returns>
        public static string ToExcelRange(this Point p)
        {
            return $"{GetExcelColumnName(p.X + 1)}{p.Y + 1}";
        }

        /// <summary>
        /// convert a rectangle to a excel range of cells
        /// </summary>
        /// <param name="r">the rectangle</param>
        /// <returns>the excel cell range (START:END)</returns>
        public static string ToExcelRange(this Rectangle r)
        {
            return $"{r.Location.ToExcelRange()}:{r.EndLocation().ToExcelRange()}";
        }

        /// <summary>
        /// convert a color to excel vba RGB() statement
        /// </summary>
        /// <param name="c">the color</param>
        /// <returns>vba RGB statement</returns>
        public static string ToExcelRGB(this Color c)
        {
            return $"RGB({c.R},{c.G},{c.B})";
        }
    }
}
