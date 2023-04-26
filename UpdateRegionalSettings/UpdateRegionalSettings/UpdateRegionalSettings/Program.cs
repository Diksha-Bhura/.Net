// See https://aka.ms/new-console-template for more information
using Microsoft.SharePoint.Client;
using PnP.Framework;
using System.Security;

Console.WriteLine("Hello, World!");
Web currentWeb;

// Starting with ClientContext, the constructor requires a URL to the
// server running SharePoint.
string userLogin = "administrator@bsdevdb.onmicrosoft.com";  //enter your account
string userPassword = "bspl@2021";  //enter your password for account
string sharePointUrl = "https://bsdevdb.sharepoint.com/sites/developersite";  // enter your site URL here

var securePassword = new SecureString();
foreach (char c in userPassword)
{
    securePassword.AppendChar(c);
}

AuthenticationManager auth = new AuthenticationManager(userLogin, securePassword);

ClientContext ctx = await auth.GetContextAsync(sharePointUrl);

    currentWeb = ctx.Web;
    ctx.Load(currentWeb);
    await ctx.ExecuteQueryRetryAsync();


//ClientContext context = await new PnP.Framework.AuthenticationManager().GetACSAppOnlyContext("https://pc45.sharepoint.com/sites/MSFT", "d6c9fb35-bfd3-4a4b-a7a8-3f565a6c9839", "NWR8Q~SW4wMt_GuZOhmDMZECzjv2YA0qaP-NEdy4");
////ClientContext context = new ClientContext("https://pc45.sharepoint.com/sites/MSFT");

//// The SharePoint web at the URL.
//Web web = context.Web;

//// We want to retrieve the web's properties.
//context.Load(web);

//// Execute the query to the server.
//context.ExecuteQuery();

// Now, the web's properties are available and we could display
// web properties, such as title.
//label1.Text = web.Title;
updateRegionalSettings(ctx, currentWeb);


void updateRegionalSettings(ClientContext context, Web web)
{
    try
    {
        web.RegionalSettings.LocaleId = 1031;
        web.RegionalSettings.ShowWeeks = true;
        web.RegionalSettings.Time24 = false;

        web.RegionalSettings.Update();
        web.Context.ExecuteQuery();
        Console.WriteLine(web.RegionalSettings.LocaleId);
    }
   catch(Exception ex) 
    {
        Console.WriteLine(ex.Message);
    }
}