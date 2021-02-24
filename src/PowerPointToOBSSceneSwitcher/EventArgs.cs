using System;

namespace PowerPointToOBSSceneSwitcher
{
   public class SceneChangedEventArgs : EventArgs
   {
      public string CurrentScene { get; set; }
   }

   public class ConnectionStateEventArgs : EventArgs
   {
      public bool IsConnected { get; set; }
   }

   public class StreamingStateEventArgs : EventArgs
   {
      public bool IsStreaming { get; set; }
   }

   public class RecordingStateEventArgs : EventArgs
   {
      public bool IsRecording { get; set; }
   }
}
