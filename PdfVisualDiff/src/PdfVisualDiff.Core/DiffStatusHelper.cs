namespace PdfVisualDiff.Core;

internal static class DiffStatusHelper
{
    public static DiffStatus FromMetrics(
        long aePixels,
        double aePercent,
        double psnr,
        double ssim,
        double dssim,
        QualityThresholds thresholds)
    {
        if (IsPass(aePixels, psnr, ssim, dssim, thresholds))
            return DiffStatus.Pass;

        return (aePercent <= thresholds.WarningAePercent || ssim >= thresholds.WarningSsim)
            ? DiffStatus.Warning
            : DiffStatus.Fail;
    }

    public static DiffStatus Merge(DiffStatus left, DiffStatus right) =>
        (DiffStatus)Math.Max((int)left, (int)right);

    private static bool IsPass(long aePixels, double psnr, double ssim, double dssim, QualityThresholds thresholds)
    {
        if (aePixels != 0)
            return false;
        if (!double.IsPositiveInfinity(psnr))
            return false;

        var ssimDelta = Math.Abs(1 - ssim);
        return ssimDelta <= thresholds.PassSsimTolerance && dssim <= thresholds.PassDssimTolerance;
    }
}
