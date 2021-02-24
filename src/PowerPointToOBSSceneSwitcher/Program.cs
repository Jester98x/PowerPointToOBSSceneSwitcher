using System;
using Microsoft.Extensions.Configuration;
using Microsoft.Office.Interop.PowerPoint;
using OBSWebsocketDotNet;
using Serilog;
//Thanks to CSharpFritz and EngstromJimmy for their gists, snippets, and thoughts.

namespace PowerPointToOBSSceneSwitcher
{
#pragma warning disable RCS1102 // Make class static.
   internal class Program
#pragma warning restore RCS1102 // Make class static.
   {
      private static readonly Application _ppt = new();
      private static readonly OBSWebsocket _obs = new();
      private static IController _controller;

      private static void Main(string[] args)
      {
         SetupStaticLogger();

         if (args.Length > 0 && args[0].Equals("OBS", StringComparison.OrdinalIgnoreCase))
         {
            Log.Information("Asking for OBS to run as the controller");
            _controller = new ObsController(_obs, _ppt);
         }
         else
         {
            Log.Information("Asking PowerPoint to run as the controller");
            _controller = new PptController(_obs, _ppt);
         }

         Console.ReadLine();
      }

      private static void SetupStaticLogger()
      {
         var configuration = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json")
            .Build();

         Log.Logger = new LoggerConfiguration()
            .ReadFrom.Configuration(configuration)
            .CreateLogger();
      }
   }
}