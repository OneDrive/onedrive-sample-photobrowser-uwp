// ------------------------------------------------------------------------------
//  Copyright (c) 2015 Microsoft Corporation
// 
//  Permission is hereby granted, free of charge, to any person obtaining a copy
//  of this software and associated documentation files (the "Software"), to deal
//  in the Software without restriction, including without limitation the rights
//  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//  copies of the Software, and to permit persons to whom the Software is
//  furnished to do so, subject to the following conditions:
// 
//  The above copyright notice and this permission notice shall be included in
//  all copies or substantial portions of the Software.
// 
//  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
//  THE SOFTWARE.
// ------------------------------------------------------------------------------

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238


namespace OneDrivePhotoBrowser
{
    using Microsoft.Graph;
    using Microsoft.OneDrive.Sdk;
    using Microsoft.OneDrive.Sdk.Authentication;
    using Models;
    using System;
    using System.Diagnostics;
    using System.Threading.Tasks;
    using Windows.UI.Xaml;
    using Windows.UI.Xaml.Controls;
    
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class AccountSelection : Page
    {
        private enum ClientType
        {
            Business,
            Consumer,
            ConsumerUwp
        }
        
        // Set these values to your app's ID and return URL.
        private readonly string oneDriveForBusinessClientId = "<Insert your OneDrive for Business client id>";
        private readonly string oneDriveForBusinessReturnUrl = "http://localhost:8080";
        private readonly string oneDriveForBusinessBaseUrl = "https://graph.microsoft.com/";

        private readonly string oneDriveConsumerClientId = "<Insert your OneDrive Consumer client id>";
        private readonly string oneDriveConsumerReturnUrl = "<Insert your OneDrive Consumer client Redirect id>";
        private readonly string oneDriveConsumerBaseUrl = "https://api.onedrive.com/v1.0";
        private readonly string[] scopes = new string[] { "onedrive.readonly", "wl.signin", "offline_access" };

        public AccountSelection()
        {
            this.InitializeComponent();
            this.Loaded += AccountSelection_Loaded;
        }

        private async void AccountSelection_Loaded(object sender, RoutedEventArgs e)
        {
            var app = ((App) Application.Current);
            if (app.OneDriveClient != null)
            {
                var msaAuthProvider = app.AuthProvider as MsaAuthenticationProvider;
                var adalAuthProvider = app.AuthProvider as AdalAuthenticationProvider;
                if (msaAuthProvider != null)
                {
                    await msaAuthProvider.SignOutAsync();
                }
                else if (adalAuthProvider != null)
                {
                    await adalAuthProvider.SignOutAsync();
                }
                
                app.OneDriveClient = null;
            }

            // Don't show AAD login if the required AAD auth values aren't set
            if (string.IsNullOrEmpty(this.oneDriveForBusinessClientId) || string.IsNullOrEmpty(this.oneDriveForBusinessReturnUrl))
            {
                this.AadButton.Visibility = Visibility.Collapsed;
            }
        }

        private void AadButton_Click(object sender, RoutedEventArgs e)
        {
            this.InitializeClient(ClientType.Business, e);
        }

        private void MsaButton_Click(object sender, RoutedEventArgs e)
        {
            this.InitializeClient(ClientType.Consumer, e);
        }

        private void OnlineId_Click(object sender, RoutedEventArgs e)
        {
            this.InitializeClient(ClientType.ConsumerUwp, e);
        }

        private async void InitializeClient(ClientType clientType, RoutedEventArgs e)
        {
            var app = (App) Application.Current;
            if (app.OneDriveClient == null)
            {
                Task authTask;

                if (clientType == ClientType.Business)
                {
                    var adalAuthProvider = new AdalAuthenticationProvider(
                        this.oneDriveForBusinessClientId,
                        this.oneDriveForBusinessReturnUrl);
                    authTask = adalAuthProvider.AuthenticateUserAsync(this.oneDriveForBusinessBaseUrl);
                    app.OneDriveClient = new OneDriveClient(this.oneDriveForBusinessBaseUrl + "/_api/v2.0", adalAuthProvider);
                    app.AuthProvider = adalAuthProvider;
                }
                else if (clientType == ClientType.ConsumerUwp)
                {
                    var onlineIdAuthProvider = new OnlineIdAuthenticationProvider(
                        this.scopes);
                    authTask = onlineIdAuthProvider.RestoreMostRecentFromCacheOrAuthenticateUserAsync();
                    app.OneDriveClient = new OneDriveClient(this.oneDriveConsumerBaseUrl, onlineIdAuthProvider);
                    app.AuthProvider = onlineIdAuthProvider;
                }
                else
                {
                    var msaAuthProvider = new MsaAuthenticationProvider(
                        this.oneDriveConsumerClientId,
                        this.oneDriveConsumerReturnUrl,
                        this.scopes,
                        new CredentialVault(this.oneDriveConsumerClientId));
                    authTask = msaAuthProvider.RestoreMostRecentFromCacheOrAuthenticateUserAsync();
                    app.OneDriveClient = new OneDriveClient(this.oneDriveConsumerBaseUrl, msaAuthProvider);
                    app.AuthProvider = msaAuthProvider;
                }

                try
                {
                    await authTask;
                    app.NavigationStack.Add(new ItemModel(new Item()));
                    this.Frame.Navigate(typeof(MainPage), e);
                }
                catch (ServiceException exception)
                {
                    // Swallow the auth exception but write message for debugging.
                    Debug.WriteLine(exception.Error.Message);
                }
            }
            else
            {
                this.Frame.Navigate(typeof(MainPage), e);
            }
        }
    }
}
