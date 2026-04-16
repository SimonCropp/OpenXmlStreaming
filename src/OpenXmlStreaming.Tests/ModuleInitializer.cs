public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifyDiffPlex.Initialize(OutputType.Compact);
        VerifierSettings.UseSsimForPng(threshold: 0.999);
        VerifierSettings.InitializePlugins();
    }
}
