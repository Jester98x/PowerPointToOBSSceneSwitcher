using Microsoft.Office.Interop.PowerPoint;
using OBSWebsocketDotNet;
using Serilog;
using System;

namespace PowerPointToOBSSceneSwitcher
{
   public interface IController
   {
      ObsLocal Obs { get; }
      PowerPoint Ppt { get; }
   }

   /// <summary>
   /// OBS is the controller and will be managing the
   /// advancement of PowerPoint slides
   /// </summary>
   public class ObsController : IController, IDisposable
   {
      private bool disposedValue;

      public ObsLocal Obs { get; }
      public PowerPoint Ppt { get; }

      public ObsController(OBSWebsocket obs, Application ppt)
      {
         Obs = new(obs);
         Ppt = new(ppt);

         Obs.SceneChanged += Obs_SceneChanged;

         Obs.Connect().GetAwaiter().GetResult();
      }

      private void Obs_SceneChanged(object sender, EventArgs args)
      {
         Log.Information(
            "OBS moved to scene: {SceneName}",
            ((SceneChangedEventArgs)args).CurrentScene);

         var commands = Obs.PowerPointCommands;

         if (commands?.Count > 0)
         {
            Ppt.ProcessCommands(commands);
         }
      }

      protected virtual void Dispose(bool disposing)
      {
         if (!disposedValue)
         {
            if (disposing)
            {
               Obs.SceneChanged -= Obs_SceneChanged;
            }

            disposedValue = true;
         }
      }

      public void Dispose()
      {
         Dispose(disposing: true);
         GC.SuppressFinalize(this);
      }
   }
}
