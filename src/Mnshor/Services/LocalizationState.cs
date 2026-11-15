namespace Mnshor.Services;

public class LocalizationState
{
    public event Action? OnChange;

    public const string DefaultCulture = "en";
    public const string SecondCulture = "ar";
    public static readonly string[] SupportedCultures = [DefaultCulture, SecondCulture];

    public string CurrentLanguage { get; private set; } = DefaultCulture;

    public bool IsRtl => CurrentLanguage == SecondCulture;
    public string Direction => IsRtl ? "rtl" : "ltr";

    public void SetLanguage(string lang)
    {
        var normalized = string.Equals(lang, SecondCulture, StringComparison.OrdinalIgnoreCase)
            ? SecondCulture
            : DefaultCulture;

        if (CurrentLanguage != normalized)
        {
            CurrentLanguage = normalized;
            OnChange?.Invoke();
        }
    }

    public void ToggleLanguage()
    {
        SetLanguage(CurrentLanguage == DefaultCulture ? SecondCulture : DefaultCulture);
    }

    public string this[string key] => TranslationDictionary.Get(key, CurrentLanguage);
}
