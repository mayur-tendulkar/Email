using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using Xamarin.Forms;

namespace Mail
{
	public partial class App : Application
	{
        public static PublicClientApplication IdentityClientApp = null;

        public static string ClientID = "36df6350-62d8-45f1-a66b-dcea0d12f768";

        public static string[] Scopes = { "User.Read", "Calendars.Read ", "Calendars.ReadWrite", "Mail.Send" };

        public static UIParent UiParent = null;

        public static DirectoryObject Me { get; set; }
        public App ()
		{
			InitializeComponent();
            IdentityClientApp = new PublicClientApplication(ClientID);
            MainPage = new MainPage();
		}

		protected override void OnStart ()
		{
			// Handle when your app starts
		}

		protected override void OnSleep ()
		{
			// Handle when your app sleeps
		}

		protected override void OnResume ()
		{
			// Handle when your app resumes
		}
	}
}
