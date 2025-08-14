using System.Text;
using DeBron.PowerPoint.Builder.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace DeBron.PowerPoint.Builder;

public static class BibletextProvider
{
    public static List<Verse> Provide(string biblebook, int chapter, int startVerse = 1, int endVerse = 1, string book = "NBV21")
    {
        var options = new ChromeOptions();
        options.AddArgument("--headless"); // Headless modus

        using var driver = new ChromeDriver(options);
        var url = $"https://www.debijbel.nl/bijbel/{book}/{biblebook}.{chapter}";
        driver.Navigate().GoToUrl(url);
        
        var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        wait.Until(d => d.FindElement(By.CssSelector(".verse")));
        
        var result = new List<Verse>();

        for (var i = startVerse; i <= endVerse; i++)
        {
            var subVerses = driver.FindElements(By.Id($"{book}.{biblebook}.{chapter}.{i}"));
            var text = string.Join("", subVerses.Select(verse => verse.Text));
            result.Add(new Verse(i, text));
        }

        driver.Quit();

        return result;
    }
}