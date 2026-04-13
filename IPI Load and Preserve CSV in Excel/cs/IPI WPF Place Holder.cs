using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

public class WatermarkAdorner : Adorner
{
    private readonly string _watermark;

    public WatermarkAdorner(UIElement adornedElement, string watermark) : base(adornedElement)
    {
        IsHitTestVisible = false;
        _watermark = watermark;
    }

    protected override void OnRender(DrawingContext drawingContext)
    {
        bool isEmpty = false;
        Control control = AdornedElement as Control;

        // Use traditional casting instead of 'is Type variable'
        TextBox tb = AdornedElement as TextBox;
        PasswordBox pb = AdornedElement as PasswordBox;
        
        if (tb != null) {
            isEmpty = string.IsNullOrEmpty(tb.Text);
        } else if (pb != null) {
            isEmpty = string.IsNullOrEmpty(pb.Password);
        }

        if (isEmpty && control != null)
        {
            var dpi = VisualTreeHelper.GetDpi(this);

            var formattedText = new FormattedText (
                _watermark,
                System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface(control.FontFamily, control.FontStyle, control.FontWeight, control.FontStretch),
                control.FontSize,
                Brushes.Gray,
                dpi.PixelsPerDip
            );

            drawingContext.DrawText(formattedText, new Point(5, 2));
        }
    }
}

public static class WatermarkService 
{
    public static readonly DependencyProperty WatermarkProperty = 
        DependencyProperty.RegisterAttached(
            "Watermark",
            typeof(string),
            typeof(WatermarkService),
            new PropertyMetadata("", OnWatermarkChanged));

    public static void SetWatermark(DependencyObject element, string value)
    {
        element.SetValue(WatermarkProperty, value);
    }

    public static string GetWatermark(DependencyObject element)
    {
        return (string)element.GetValue(WatermarkProperty);
    }

    private static void OnWatermarkChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        Control control = d as Control;
        if (control == null) return;

        // Use traditional casting instead of 'is Type variable'
        TextBox tb = d as TextBox;
        PasswordBox pb = d as PasswordBox;

        control.Loaded += (s, ev) => {
            var layer = AdornerLayer.GetAdornerLayer(control);
            if (layer != null)
            {
                layer.Add(new WatermarkAdorner(control, GetWatermark(control)));
            }
        };

        if (tb != null)
        {
            tb.TextChanged += (s, ev) =>
            {
                var layer = AdornerLayer.GetAdornerLayer(tb);
                if (layer != null) {
                    layer.Update(tb);
                }
            };
        }
        else if (pb != null) {
            // PasswordBox uses PassWordChanged instead of TextChanged
            pb.PasswordChanged += (s, ev) =>
            {
                var layer = AdornerLayer.GetAdornerLayer(pb);
                if (layer != null) {
                    layer.Update(pb);
                }
            };
        }
    }
}