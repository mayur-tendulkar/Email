using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;

namespace Mail
{
	public partial class MainPage : ContentPage
	{
        private static GraphServiceClient Client;
        private User Me;

		public MainPage()
		{
			InitializeComponent();
		}

        protected async override void OnAppearing()
        {
            base.OnAppearing();
            if (App.IdentityClientApp.Users.Any())
            {
                Client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                               new DelegateAuthenticationProvider(
                                   async (requestMessage) =>
                                   {
                                       var tokenRequest = await App.IdentityClientApp.AcquireTokenSilentAsync(App.Scopes, App.IdentityClientApp.Users.FirstOrDefault());
                                       requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                                   }));
                Me = await Client.Me.Request().GetAsync();
                Username.Text = $"Welcome {((User)Me).DisplayName}";
            }
        }

        private async void AuthenticateClicked(object sender, EventArgs e)
        {
            try
            {
                Client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                      new DelegateAuthenticationProvider(async(requestMessage) =>
                {
                    var tokenRequest = await App.IdentityClientApp.AcquireTokenAsync(App.Scopes, App.UiParent).ConfigureAwait(false);
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                }));
                Me = await Client.Me.Request().GetAsync();
                Username.Text = $"Welcome {((User)Me).DisplayName}";
            }
            catch (MsalException ex)
            {
                await DisplayAlert("Error", ex.Message, "OK", "Cancel");
            }

        }

        private async void ApplyLeaveClicked(object sender, EventArgs e)
        {
            Client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        var tokenRequest = await App.IdentityClientApp.AcquireTokenSilentAsync(App.Scopes, App.IdentityClientApp.Users.FirstOrDefault());
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                    }));
            var message = new Message();
            var to = new List<Recipient>();
            var recipient = new Recipient
            {
                EmailAddress = new EmailAddress() { Name = "My Manager", Address = "{valid email address}" }
            };
            to.Add(recipient);
            message.ToRecipients = to;
            message.Body = new ItemBody() { Content = "I will be taking leave today.", ContentType = BodyType.Text };
            message.Subject = "[Demo] Leave Application";
            await Client.Me.SendMail(message).Request().PostAsync();
          
        }
    }
}
