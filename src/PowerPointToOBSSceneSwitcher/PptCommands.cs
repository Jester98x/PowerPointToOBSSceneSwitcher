namespace PowerPointToOBSSceneSwitcher
{

   public class PptCommands
   {
      public enum CommandType
      {
         NextSlide,
         PreviousSlide,
         SetText,
         ClickSlide,
         ClickButton,
      }

      public CommandType Command { get; }

      public PptCommands(CommandType command)
      {
         Command = command;
      }
   }
}
