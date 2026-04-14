public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifyDiffPlex.Initialize(OutputType.Compact);
        VerifyImageSharp.Initialize(ssimThreshold: 0.999);
        VerifierSettings.InitializePlugins();
    }
}
