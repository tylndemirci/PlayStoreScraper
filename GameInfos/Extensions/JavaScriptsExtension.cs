using OpenQA.Selenium;

namespace GameInfos.Extensions
{
    public static class JavaScriptsExtension
    {
        public static IJavaScriptExecutor Scripts(this IWebDriver driver)
        {
            return (IJavaScriptExecutor)driver;
        }
    }
}