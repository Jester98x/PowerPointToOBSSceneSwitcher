using Microsoft.Office.Interop.PowerPoint;
using OBSWebsocketDotNet;
using Serilog;
using System;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace PowerPointToOBSSceneSwitcher
{
   /// <summary>
   /// PowerPoint is the controller and well be manageing
   /// the selection of scenes within OBS
   /// </summary>
   public partial class PptController : IController, IDisposable
   {
      private bool _canSwitchScene = true;
      private bool _isConnected;
      private readonly System.Threading.Timer _connectionTimer;
      private bool disposedValue;

      public ObsLocal Obs { get; }
      public PowerPoint Ppt { get; }

      public PptController(OBSWebsocket obs, Application ppt)
      {
         Obs = new(obs);
         Ppt = new(ppt);

         Ppt.Application.SlideShowNextSlide += App_SlideShowNextSlide;
         _connectionTimer = new System.Threading.Timer(
            ConnectionTimerCallback,
            null,
            TimeSpan.FromSeconds(5),
            TimeSpan.FromSeconds(5));

         InitObs();
      }

      private void Obs_ConnectionStateChanged(object sender, EventArgs e)
      {
         _isConnected = ((ConnectionStateEventArgs)e).IsConnected;

         HandleDisconnected();
         HandleConnected();
      }

      private void Obs_StreamingStateChanged(object sender, EventArgs e)
      {
         // Do something to indicate that OBS is streaming or not
         if (((StreamingStateEventArgs)e).IsStreaming)
         {
            // We're stremaing !!!
         }
         else
         {
            // We're not streaming !!!
         }
      }

      private void Obs_RecordingStateChanged(object sender, EventArgs e)
      {
         // Do somting to indicate that OBS is recording or not
         if (((RecordingStateEventArgs)e).IsRecording)
         {
            // We're recording !!!
         }
         else
         {
            // We're not recording !!!
         }
      }

      private void InitObs()
      {
         try
         {
            Obs.ConnectionStateChanged += Obs_ConnectionStateChanged;
            Obs.RecordingStateChanged += Obs_RecordingStateChanged;
            Obs.StreamingStateChanged += Obs_StreamingStateChanged;

            Obs.Connect().GetAwaiter().GetResult();
         }
         catch
         {
            // Unable to connect to OBS
         }
      }

      private void ConnectionTimerCallback(object state)
      {
         HandleDisconnected();
      }

      private void HandleDisconnected()
      {
         if (_isConnected)
         {
            return;
         }

         Obs.Connect().GetAwaiter().GetResult();
         _connectionTimer.Change(TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(5));
      }

      private void HandleConnected()
      {
         if (!_isConnected)
         {
            return;
         }

         _connectionTimer.Change(-1, -1);
      }

      public void App_SlideShowNextSlide(SlideShowWindow Wn)
      {
         if (Wn == null)
         {
            return;
         }

         Log.Information("Moved to Slide Number {SlideNumber}", Wn.View.Slide.SlideNumber);

         SwitchScene(Wn).GetAwaiter().GetResult();
      }

      private async Task SwitchScene(SlideShowWindow Wn)
      {
         var sceneChanged = false;

         string[] obsCommands = default;
         try
         {
            obsCommands = Wn.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text.Split('\r');
         }
         catch
         {
            // Nothing to read
         }

         if (obsCommands.Length == 0)
         {
            return;
         }

         HandleNoSceneSwitch(obsCommands);

         foreach (var obsCommand in obsCommands)
         {
            if (obsCommand.StartsWith("OBSScene:", StringComparison.OrdinalIgnoreCase))
            {
               var obsSceneName = obsCommand.Substring(9).Trim();

               _ = await Obs.ChangeScene(obsSceneName);

               sceneChanged = true;
            }

            if (obsCommand.StartsWith("OBSDelay:", StringComparison.OrdinalIgnoreCase))
            {
               var delay = Convert.ToInt32(obsCommand[9..].Trim());
               await Task.Delay(delay);
            }
            else if (obsCommand.StartsWith("OBSHotKeys:", StringComparison.OrdinalIgnoreCase))
            {
               var (keyName, modifiers) = ParseHotKeys(obsCommand[11..].Trim());
               await Obs.SendHotKeys(keyName, modifiers);
            }
            else if (obsCommand.StartsWith("OBSDefault:", StringComparison.OrdinalIgnoreCase))
            {
               Obs.DefaultScene = obsCommand[11..].Trim();
            }
            else if (obsCommand.StartsWith("OBSRecord:", StringComparison.OrdinalIgnoreCase))
            {
               _ = await Obs.StartStopRecording(obsCommand[10..].Trim());
            }
            else if (obsCommand.StartsWith("OBSStream:", StringComparison.OrdinalIgnoreCase))
            {
               _ = await Obs.StartStopStreaming(obsCommand[10..].Trim());
            }
         }

         if (_canSwitchScene && !sceneChanged)
         {
            _ = await Obs.GotoDefault();
         }

         _canSwitchScene = true;
      }

      private (string keyName, char[] modifiers) ParseHotKeys(string value)
      {
         var parts = value.Split('|', StringSplitOptions.RemoveEmptyEntries);

         if (parts.Length > 1)
         {
            return (parts[1], parts[0].ToCharArray());
         }

         return (parts[0], "".ToCharArray());
      }

      private void HandleNoSceneSwitch(string[] obsCommands)
      {
         if (obsCommands.Contains("OBSStay", StringComparer.OrdinalIgnoreCase))
         {
            _canSwitchScene = false;
         }
      }

      protected virtual void Dispose(bool disposing)
      {
         if (!disposedValue)
         {
            if (disposing)
            {
               Ppt.Application.SlideShowNextSlide -= App_SlideShowNextSlide;
               Obs.ConnectionStateChanged -= Obs_ConnectionStateChanged;
               Obs.RecordingStateChanged -= Obs_RecordingStateChanged;
               Obs.StreamingStateChanged -= Obs_StreamingStateChanged;
               Obs.Dispose();
            }

            disposedValue = true;
         }
      }

      public void Dispose()
      {
         Log.Information("Shutting down controller");
         Dispose(disposing: true);
         GC.SuppressFinalize(this);
      }
   }
}
