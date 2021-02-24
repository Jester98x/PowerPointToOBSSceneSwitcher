using OBSWebsocketDotNet;
using OBSWebsocketDotNet.Types;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PowerPointToOBSSceneSwitcher
{
   public class ObsLocal : IDisposable
   {
      private bool _isConnected;
      private bool _isStreaming;
      private bool _isRecording;
      private readonly Dictionary<string, List<SceneItem>> _validScenes = new();
      private readonly Dictionary<string, TextGDIPlusProperties> _pptCommands = new();
      private string _defaultScene;
      private string _currentScene;
      private bool disposedValue;

      public OBSWebsocket Application { get; }

      public List<PptCommands> PowerPointCommands
      {
         get
         {
            _pptCommands.TryGetValue(_currentScene, out var c);
            if (c != null)
            {
               Log.Debug("ppt_commands are {@Commands}", c);
               return ParseCommands(c.Text);
            }

            return new();
         }
      }

      private List<PptCommands> ParseCommands(string rawCommands)
      {
         // Each command should appear on it's own line
         if (string.IsNullOrWhiteSpace(rawCommands))
         {
            return new();
         }

         var commands = rawCommands.Split("\n");

         List<PptCommands> pptCommands = new();
         foreach (var command in commands)
         {
            if (command.Equals("next_slide", StringComparison.OrdinalIgnoreCase))
            {
               pptCommands.Add(new(PptCommands.CommandType.NextSlide));
            }
            else if (command.Equals("previous_slide", StringComparison.OrdinalIgnoreCase))
            {
               pptCommands.Add(new(PptCommands.CommandType.PreviousSlide));
            }
            else if (command.Equals("click_slide", StringComparison.OrdinalIgnoreCase))
            {
               pptCommands.Add(new(PptCommands.CommandType.ClickSlide));
            }
         }

         return pptCommands;
      }

      public string DefaultScene
      {
         get => _defaultScene;
         set
         {
            if (_validScenes?.ContainsKey(value) == true)
            {
               _defaultScene = value;
            }
         }
      }

      public event EventHandler StreamingStateChanged;
      public event EventHandler RecordingStateChanged;
      public event EventHandler ConnectionStateChanged;
      public event EventHandler SceneChanged;

      protected virtual void OnSceneChanged(SceneChangedEventArgs e)
      {
         var handler = SceneChanged;
         e.CurrentScene = _currentScene;
         handler?.Invoke(this, e);
      }

      protected virtual void OnConnectionStateChanged(ConnectionStateEventArgs e)
      {
         var handler = ConnectionStateChanged;
         e.IsConnected = _isConnected;
         handler?.Invoke(this, e);
      }

      protected virtual void OnStreamingStateChanged(StreamingStateEventArgs e)
      {
         var handler = StreamingStateChanged;
         e.IsStreaming = _isStreaming;
         handler?.Invoke(this, e);
      }

      protected virtual void OnRecordingStateChanged(RecordingStateEventArgs e)
      {
         var handler = RecordingStateChanged;
         e.IsRecording = _isRecording;
         handler?.Invoke(this, e);
      }

      public ObsLocal(OBSWebsocket app)
      {
         Application = app;
      }

      public async Task Connect()
      {
         Application.Connected += Obs_Connected;
         Application.Disconnected += Obs_Disconnected;

         Application.SceneCollectionChanged += Obs_SceneCollectionChanged;
         Application.SceneListChanged += Obs_SceneListChanged;
         Application.SceneChanged += Obs_SceneChanged;
         Application.RecordingStateChanged += Obs_RecordingStateChanged;
         Application.StreamingStateChanged += Obs_StreamingStateChanged;
         await Application.Connect("ws://127.0.0.1:4444", string.Empty);
      }

      private void Obs_SceneChanged(OBSWebsocket sender, string newSceneName)
      {
         _currentScene = newSceneName;
         OnSceneChanged(new SceneChangedEventArgs
         {
            CurrentScene = _currentScene
         });
         Log.Information("OBS reports scene changed to {SceneName}", newSceneName);
      }

      private void Obs_SceneListChanged(object sender, EventArgs e)
      {
         GetScenes().GetAwaiter().GetResult();
      }

      private void Obs_SceneCollectionChanged(object sender, EventArgs e)
      {
         GetScenes().GetAwaiter().GetResult();
      }

      private void Obs_StreamingStateChanged(OBSWebsocket sender, OutputState newState)
      {
         if (newState == OutputState.Started)
         {
            _isStreaming = true;
         }
         else if (newState == OutputState.Stopped)
         {
            _isStreaming = false;
         }

         Log.Information("OBS reports streaming state is now {CurrentState}", newState);
      }

      private void Obs_RecordingStateChanged(OBSWebsocket sender, OutputState newState)
      {
         if (newState == OutputState.Started)
         {
            _isRecording = true;
         }
         else if (newState == OutputState.Stopped)
         {
            _isRecording = false;
         }

         Log.Information("OBS reports recording state is now {CurrentState}", newState);
      }

      private void Obs_Disconnected(object sender, EventArgs e)
      {
         _isConnected = false;
         OnConnectionStateChanged(new ConnectionStateEventArgs
         {
            IsConnected = _isConnected
         });

         Log.Information("Disconnected from OBS");
      }

      private void Obs_Connected(object sender, EventArgs e)
      {
         _isConnected = true;
         OnConnectionStateChanged(new ConnectionStateEventArgs
         {
            IsConnected = _isConnected
         });

         Log.Information("Connected to OBS");

         GetScenes().GetAwaiter().GetResult();
         GetPPtCommands().GetAwaiter().GetResult();
      }

      private async Task GetPPtCommands()
      {
         if (!_isConnected)
         {
            return;
         }

         foreach (var scene in _validScenes)
         {
            var commands = scene.Value.FirstOrDefault(s =>
               s.SourceName.StartsWith("ppt_commands", StringComparison.OrdinalIgnoreCase)
               && s.InternalType.Equals("text_gdiplus_v2", StringComparison.OrdinalIgnoreCase));

            if (commands != null)
            {
               var ppt_command = await Application.GetTextGDIPlusProperties(commands.SourceName);
               _pptCommands.TryAdd(scene.Key, ppt_command);
            }
         }
      }

      public async Task GetScenes()
      {
         if (!_isConnected)
         {
            return;
         }

         var allScenes = await Application.GetSceneList();
         foreach (var scene in allScenes.Scenes)
         {
            _validScenes.TryAdd(scene.Name, scene.Items);
         }

         Log.Information("OBS reports currents scenes are {@SceneList}", _validScenes);
      }

      public async Task<bool> ChangeScene(string scene)
      {
         if (!_isConnected)
         {
            return false;
         }

         if (!_validScenes.ContainsKey(scene))
         {
            if (string.IsNullOrEmpty(_defaultScene))
            {
               return false;
            }

            scene = _defaultScene;
         }

         await Application.SetCurrentScene(scene);
         _currentScene = scene;

         return true;
      }

      public async Task SendHotKeys(string key, params char[] keyModifiers)
      {
         if (!_isConnected)
         {
            return;
         }

         if (!Enum.TryParse(key, out OBSHotkey obsKey))
         {
            obsKey = OBSHotkey.OBS_KEY_NONE;
         }

         var obsModifiers = KeyModifier.None;
         if (keyModifiers.Length > 0)
         {
            for (var i = 0; i < keyModifiers.Length; i++)
            {
               switch (keyModifiers[i])
               {
                  case '+':
                     obsModifiers |= KeyModifier.Shift;
                     break;
                  case '^':
                     obsModifiers |= KeyModifier.Control;
                     break;
                  case '%':
                     obsModifiers |= KeyModifier.Alt;
                     break;
               }
            }
         }

         await Application.TriggerHotkeyBySequence(obsKey, obsModifiers);
      }

      public async Task SendHotKeysByName(string hotkeyName)
      {
         if (!_isConnected)
         {
            return;
         }

         await Application.TriggerHotkeyByName(hotkeyName);
      }

      public async Task<bool> StartStopRecording(string command)
      {
         if (!_isConnected)
         {
            return false;
         }

         try
         {
            if (command.Equals("start", StringComparison.OrdinalIgnoreCase))
            {
               await Application.StartRecording();
               _isRecording = true;
            }
            else if (command.Equals("stop", StringComparison.OrdinalIgnoreCase))
            {
               await Application.StopRecording();
               _isRecording = false;
            }

            OnRecordingStateChanged(new RecordingStateEventArgs
            {
               IsRecording = _isRecording
            });
         }
         catch
         {
            // Already doing it
         }

         return true;
      }

      public async Task<bool> StartStopStreaming(string command)
      {
         if (!_isConnected)
         {
            return false;
         }

         try
         {
            if (command.Equals("start", StringComparison.OrdinalIgnoreCase))
            {
               await Application.StartStreaming();
               _isStreaming = true;
            }
            else if (command.Equals("stop", StringComparison.OrdinalIgnoreCase))
            {
               await Application.StopStreaming();
               _isStreaming = false;
            }

            OnStreamingStateChanged(new StreamingStateEventArgs
            {
               IsStreaming = _isStreaming
            });
         }
         catch
         {
            // Already doing it
         }

         return true;
      }

      public async Task<bool> GotoDefault()
      {
         if (string.IsNullOrWhiteSpace(_defaultScene))
         {
            return false;
         }

         Log.Information("Requesting OBS go to default scene {SceneName}", _defaultScene);

         return await ChangeScene(_defaultScene);
      }

      protected virtual void Dispose(bool disposing)
      {
         if (!disposedValue)
         {
            if (disposing)
            {
               Application.Connected += Obs_Connected;
               Application.Disconnected += Obs_Disconnected;

               Application.SceneCollectionChanged += Obs_SceneCollectionChanged;
               Application.SceneListChanged += Obs_SceneListChanged;
               Application.RecordingStateChanged += Obs_RecordingStateChanged;
               Application.StreamingStateChanged += Obs_StreamingStateChanged;

               Application.Disconnect();
            }

            disposedValue = true;
         }
      }

      public void Dispose()
      {
         // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
         Dispose(disposing: true);
         GC.SuppressFinalize(this);
      }
   }
}
