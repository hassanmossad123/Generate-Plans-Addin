using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Security.Policy;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace generate_floor_ceiling_plans
{
    [Transaction(TransactionMode.Manual)]
    public class Application : IExternalApplication
    {
        // ========== CONSTANT DEFINITIONS ==========
        private const string TAB_NAME = "Tools";
        private const string PANEL_NAME = "Create Plans";
        private const string BUTTON_NAME = "Add Plans";
        private const string BUTTON_ID = "Button_01";
        private const string COMMAND_CLASS_NAME = "generate_floor_ceiling_plans.Command";
        private const string TOOLTIP_TEXT = "Generate Plans";
        private const string HELP_URL = "www.google.com";
        private const string LONG_DESCRIPTION =
            "This tool will help you to create floor and ceiling plans quickly for selected levels and apply view templates to them.";

        // ========== PUBLIC METHODS ==========
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                CreateRibbonInterface(application);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ribbon Initialization Error",
                    $"Failed to initialize ribbon: {ex.Message}");
                return Result.Failed;
            }
        }

        // ========== PRIVATE METHODS ==========
        private void CreateRibbonInterface(UIControlledApplication application)
        {
            application.CreateRibbonTab(TAB_NAME);

            RibbonPanel ribbonPanel = application.CreateRibbonPanel(
                TAB_NAME,
                PANEL_NAME);

            PushButton button = CreateRibbonButton(ribbonPanel);
            ConfigureButton(button, Properties.Resources.pushButton, Properties.Resources.tooltip_description, HELP_URL);
        }

        private PushButton CreateRibbonButton(RibbonPanel ribbonPanel)
        {
            PushButtonData buttonData = new PushButtonData(
                BUTTON_ID,
                BUTTON_NAME,
                Assembly.GetExecutingAssembly().Location,
                COMMAND_CLASS_NAME);

            return ribbonPanel.AddItem(buttonData) as PushButton;
        }

        private void ConfigureButton(PushButton button, Image pushButtonImage, Image tooltipDescImage, String url)
        {
            button.LargeImage = LoadImage(pushButtonImage);
            button.ToolTip = TOOLTIP_TEXT;
            button.LongDescription = LONG_DESCRIPTION;
            button.ToolTipImage = LoadImage(tooltipDescImage);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, url));
        }

        private ImageSource LoadImage(Image image)
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            MemoryStream memoryStream = new MemoryStream();
            image.Save(memoryStream, ImageFormat.Png);
            memoryStream.Seek(0, SeekOrigin.Begin);
            bitmap.StreamSource = memoryStream;
            bitmap.EndInit();
            bitmap.Freeze();
            return bitmap;
        }
    }
}