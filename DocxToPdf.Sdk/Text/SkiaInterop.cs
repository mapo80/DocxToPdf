using SkiaSharp;
using System;
using System.Runtime.InteropServices;

namespace DocxToPdf.Sdk.Text;

internal static class SkiaInterop
{
    [DllImport("libSkiaSharp", CallingConvention = CallingConvention.Cdecl)]
    private static extern void sk_canvas_draw_text_blob(IntPtr canvas, IntPtr textBlob, float x, float y, IntPtr paint);

    public static void DrawTextBlob(SKCanvas canvas, SKTextBlob blob, float x, float y, SKPaint paint)
    {
        if (canvas == null)
            throw new ArgumentNullException(nameof(canvas));
        if (blob == null)
            throw new ArgumentNullException(nameof(blob));
        if (paint == null)
            throw new ArgumentNullException(nameof(paint));

        if (canvas.Handle == IntPtr.Zero || blob.Handle == IntPtr.Zero || paint.Handle == IntPtr.Zero)
            return;

        sk_canvas_draw_text_blob(canvas.Handle, blob.Handle, x, y, paint.Handle);
    }
}
