using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Security.Claims;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.Extensions.Options;

namespace DeBron.PowerPoint.Builder.Components.Pages;

public class PincodeModel : PageModel
{
    private readonly AuthenticationOptions _authenticationOptions;

    public PincodeModel(IOptions<AuthenticationOptions> environmentVariables)
    {
        _authenticationOptions = environmentVariables.Value;
    }

    public IActionResult OnGet(string pin)
    {
        if (pin == _authenticationOptions.Pincode)
        {
            var claims = new List<Claim>
            {
                new Claim(ClaimTypes.Name, "PincodeGebruiker")
            };

            var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
            var principal = new ClaimsPrincipal(identity);

            HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal).Wait();
            
            return LocalRedirect("~/");
        }

        return LocalRedirect("~/login?error=true");
    }
}
