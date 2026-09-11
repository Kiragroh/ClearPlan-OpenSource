using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using ClearPlan.Core.Integration;

namespace ClearPlan.Core.Tests
{
    internal static class AriaOAuthTests
    {
        public static void Contract()
        {
            TestAssert.True(typeof(AriaReportDocumentBuilder).Assembly.GetType("ClearPlan.Core.Integration.AriaOAuthTokenProvider")!=null,
                "ARIA OAuth must have a bounded, non-logging token provider.");
        }
        public static void Protocol()
        {
            int secrets=0;
            var handler=new Handler(async request=> {
                TestAssert.Equal("https://example.invalid/token",request.RequestUri.AbsoluteUri);
                TestAssert.Equal(HttpMethod.Post,request.Method);
                TestAssert.True(request.Headers.Authorization==null);
                var form=await request.Content.ReadAsStringAsync();
                TestAssert.True(form.Contains("grant_type=client_credentials") && form.Contains("client_secret=SYNTHETIC%3Da%23b%21"));
                return Json("{\"access_token\":\"SYNTHETIC_TOKEN\",\"token_type\":\"Bearer\",\"expires_in\":300}");
            });
            using(var http=new HttpClient(handler))
            {
                var provider=new AriaOAuthTokenProvider(http,new Uri("https://example.invalid/token"),"synthetic-client","system/Patient.rs",
                    ct=> { secrets++; return Task.FromResult("SYNTHETIC=a#b!"); });
                TestAssert.Equal("SYNTHETIC_TOKEN",provider.GetTokenAsync(CancellationToken.None).GetAwaiter().GetResult());
                TestAssert.Equal("SYNTHETIC_TOKEN",provider.GetTokenAsync(CancellationToken.None).GetAwaiter().GetResult());
                TestAssert.Equal(1,handler.Count); TestAssert.Equal(1,secrets);
            }
        }
        public static void Cancellation()
        {
            var handler=new Handler(request=>Task.FromResult(Json("{}")));
            using(var http=new HttpClient(handler)) using(var cancel=new CancellationTokenSource())
            {
                cancel.Cancel(); bool canceled=false;
                var provider=new AriaOAuthTokenProvider(http,new Uri("https://example.invalid/token"),"synthetic","system/Patient.rs",ct=>Task.FromResult("secret"));
                try {provider.GetTokenAsync(cancel.Token).GetAwaiter().GetResult();}catch(OperationCanceledException){canceled=true;}
                TestAssert.True(canceled); TestAssert.Equal(0,handler.Count);
            }
        }
        public static void Rejections()
        {
            foreach(string body in new[]{"SYNTHETIC_PRIVATE_ERROR", "{\"access_token\":\"bad\\nvalue\"}","{\"access_token\":\"x\",\"token_type\":\"Basic\"}","{}",new string('x',65537)})
            {
                var handler=new Handler(request=>Task.FromResult(Json(body)));
                using(var http=new HttpClient(handler))
                {
                    var provider=new AriaOAuthTokenProvider(http,new Uri("https://example.invalid/token"),"synthetic","system/Patient.rs",ct=>Task.FromResult("SYNTHETIC_SECRET"));
                    Reject(()=>provider.GetTokenAsync(CancellationToken.None).GetAwaiter().GetResult());
                    TestAssert.Equal(1,handler.Count);
                }
            }
            foreach(var status in new[]{HttpStatusCode.Unauthorized,HttpStatusCode.Redirect,HttpStatusCode.InternalServerError})
            {
                var handler=new Handler(request=>Task.FromResult(new HttpResponseMessage(status){Content=new StringContent("SYNTHETIC_PRIVATE_ERROR")}));
                using(var http=new HttpClient(handler))
                {
                    var provider=new AriaOAuthTokenProvider(http,new Uri("https://example.invalid/token"),"synthetic","system/Patient.rs",ct=>Task.FromResult("SYNTHETIC_SECRET"));
                    Reject(()=>provider.GetTokenAsync(CancellationToken.None).GetAwaiter().GetResult()); TestAssert.Equal(1,handler.Count);
                }
            }
        }
        private static HttpResponseMessage Json(string content)
        { return new HttpResponseMessage(HttpStatusCode.OK){Content=new StringContent(content)}; }
        private static void Reject(Action action)
        {
            bool failed=false;
            try {action();}catch(Exception error){failed=true; TestAssert.True(error.InnerException==null); TestAssert.False(error.Message.Contains("SYNTHETIC") || error.Message.Contains("example.invalid"));}
            TestAssert.True(failed,"Invalid OAuth response was accepted.");
        }
        private sealed class Handler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage,Task<HttpResponseMessage>> send;
            internal int Count;
            internal Handler(Func<HttpRequestMessage,Task<HttpResponseMessage>> action){send=action;}
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,CancellationToken cancellation)
            { Count++; return send(request); }
        }
    }
}
