using ExeclWeb.Middleware;
using Microsoft.AspNetCore.Builder;

namespace ExeclWeb.Extension
{
    public static class SheetWebSocketMiddlewareExtension
    {
        public static IApplicationBuilder SheetWebSocket(this IApplicationBuilder builder)
        {
            return builder.UseMiddleware<SheetWebSocketMiddleware>();
        }
    }
}
