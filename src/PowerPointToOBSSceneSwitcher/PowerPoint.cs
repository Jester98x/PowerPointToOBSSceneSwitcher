using Microsoft.Office.Interop.PowerPoint;
using Serilog;
using System.Collections.Generic;

namespace PowerPointToOBSSceneSwitcher
{
   public class PowerPoint
   {
      private bool _sideShowRunning;

      public Application Application { get; }

      public PowerPoint(Application app)
      {
         Application = app;
         Application.SlideShowBegin += Application_SlideShowBegin;
         Application.SlideShowEnd += Application_SlideShowEnd;
      }

      private void Application_SlideShowEnd(Presentation Pres)
      {
         _sideShowRunning = false;
      }

      private void Application_SlideShowBegin(SlideShowWindow Wn)
      {
         _sideShowRunning = true;
      }

      public void DispalyNextSlide()
      {
         if (!_sideShowRunning)
         {
            Application.ActivePresentation.SlideShowSettings.Run();
         }

         ////var slides = Application.ActivePresentation.Slides;
         ////foreach (Slide slide in slides)
         ////{
         ////   Log.Debug("Slides are {Name}, {Id}", slide.Name, slide.SlideID);
         ////}


         Log.Information("Asking PowetPoint to advance to next slide");
         Application.SlideShowWindows[1].View.Next();
      }

      public void DisplayPreviousSlide()
      {
         if (!_sideShowRunning)
         {
            Application.ActivePresentation.SlideShowSettings.Run();
         }

         Log.Information("Asking PowetPoint to retreat to previous slide");
         Application.SlideShowWindows[1].View.Previous();
      }

      public void ClickSlide()
      {
         if (!_sideShowRunning)
         {
            Application.ActivePresentation.SlideShowSettings.Run();
         }

         Log.Information("Asking PowetPoint to click on current slide");
         var ci = Application.SlideShowWindows[1].View.GetClickIndex();
         var totalClicks = Application.SlideShowWindows[1].View.GetClickCount();
         if (ci <= totalClicks)
         {
            Application.SlideShowWindows[1].View.GotoClick(ci + 1);
         }
      }

      public void ProcessCommands(List<PptCommands> commands)
      {
         if (commands == null || commands.Count == 0)
         {
            return;
         }

         foreach(var command in commands)
         {
            if (command.Command == PptCommands.CommandType.NextSlide )
            {
               DispalyNextSlide();
            }
            else if (command.Command == PptCommands.CommandType.PreviousSlide)
            {
               DisplayPreviousSlide();
            }
            else if (command.Command == PptCommands.CommandType.ClickSlide)
            {
               ClickSlide();
            }
         }
      }
   }
}
