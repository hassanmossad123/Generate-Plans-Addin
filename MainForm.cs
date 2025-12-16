using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Color = System.Drawing.Color;

namespace generate_floor_ceiling_plans
{
    public partial class MainForm : System.Windows.Forms.Form
    {
        // ========== CONSTANT DEFINITIONS ==========
        private const string ALL_LEVELS_OPTION = "All Levels";
        private const string NO_SELECTION_TITLE = "No Selection";
        private const string NO_SELECTION_MESSAGE = "Please select at least one discipline to generate.";
        private const string LEVEL_REQUIRED_TITLE = "Selection Required";
        private const string LEVEL_REQUIRED_MESSAGE = "Please select a level.";

        // ========== PRIVATE FIELDS ==========
        private UIDocument _uiDocument;
        private Document _document;
        private List<Level> _availableLevels = new List<Level>();
        private ExternalEvent _externalEvent;
        private GeneratePlanHandler _planHandler;

        // ========== CONSTRUCTOR ==========
        public MainForm(UIDocument uiDocument)
        {
            InitializeComponent();
            InitializeDefaultSelections();
            SetupRevitComponents(uiDocument);
            PopulateLevelComboBox();
        }

        // ========== EVENT HANDLERS ==========
        private void bt_generate_Click(object sender, EventArgs e)
        {
            if (!ValidateUserInput())
                return;

            if (!IsAnyDisciplineSelected())
            {
                ShowWarningMessage(NO_SELECTION_MESSAGE, NO_SELECTION_TITLE);
                return;
            }

            PrepareAndExecutePlanGeneration();
        }

        private void bt_cancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            UpdateCheckBoxAppearance(sender as CheckBox);
        }

        // ========== PRIVATE METHODS ==========
        private void PrepareAndExecutePlanGeneration()
        {
            Level selectedLevel = GetSelectedLevel();
            Dictionary<string, (bool FloorPlan, bool CeilingPlan)> disciplineSelections = GetDisciplineSelections();

            _planHandler.Initialize(selectedLevel, disciplineSelections);
            _externalEvent.Raise();

            this.Close();
        }

        private Level GetSelectedLevel()
        {
            if (cb_lvl.SelectedItem.ToString() == ALL_LEVELS_OPTION)
                return null;

            return _availableLevels[cb_lvl.SelectedIndex];
        }

        private Dictionary<string, (bool FloorPlan, bool CeilingPlan)> GetDisciplineSelections()
        {
            return new Dictionary<string, (bool FloorPlan, bool CeilingPlan)>
            {
                [GeneratePlanHandler.POWER_SUB_DISCIPLINE] = (cb_power_floor.Checked, cb_power_ceiling.Checked),
                [GeneratePlanHandler.LIGHTING_SUB_DISCIPLINE] = (cb_lighting_floor.Checked, cb_lighting_ceiling.Checked),
                [GeneratePlanHandler.HVAC_SUB_DISCIPLINE] = (cb_hvac_floor.Checked, cb_hvac_ceiling.Checked),
                [GeneratePlanHandler.FIRE_ALARM_SUB_DISCIPLINE] = (cb_fa_floor.Checked, cb_fa_ceiling.Checked),
                [GeneratePlanHandler.LIGHT_CURRENT_SUB_DISCIPLINE] = (cb_lc_floor.Checked, cb_lc_ceiling.Checked),
                [GeneratePlanHandler.COORDINATION_SUB_DISCIPLINE] = (cb_coordination_floor.Checked, cb_coordination_ceiling.Checked)
            };
        }

        private bool IsAnyDisciplineSelected()
        {
            return cb_power_floor.Checked || cb_power_ceiling.Checked ||
                   cb_lighting_floor.Checked || cb_lighting_ceiling.Checked ||
                   cb_hvac_floor.Checked || cb_hvac_ceiling.Checked ||
                   cb_fa_floor.Checked || cb_fa_ceiling.Checked ||
                   cb_lc_floor.Checked || cb_lc_ceiling.Checked ||
                   cb_coordination_floor.Checked || cb_coordination_ceiling.Checked;
        }

        private bool ValidateUserInput()
        {
            if (cb_lvl.SelectedIndex < 0)
            {
                ShowWarningMessage(LEVEL_REQUIRED_MESSAGE, LEVEL_REQUIRED_TITLE);
                return false;
            }
            return true;
        }

        private void PopulateLevelComboBox()
        {
            _availableLevels.Clear();
            cb_lvl.Items.Clear();

            List<Level> documentLevels = _planHandler.GetAllDocumentLevels();
            foreach (Level level in documentLevels.OrderBy(level => level.Elevation))
            {
                _availableLevels.Add(level);
                cb_lvl.Items.Add(level.Name);
            }

            cb_lvl.Items.Add(ALL_LEVELS_OPTION);

            if (cb_lvl.Items.Count > 0)
                cb_lvl.SelectedIndex = 0;
        }

        private void SetupRevitComponents(UIDocument uiDocument)
        {
            _uiDocument = uiDocument;
            _document = _uiDocument.Document;

            _planHandler = new GeneratePlanHandler(_document);
            _externalEvent = ExternalEvent.Create(_planHandler);
        }

        private void InitializeDefaultSelections()
        {
            // Set reasonable defaults for common scenarios
            cb_power_floor.Checked = true;
            cb_lighting_floor.Checked = true;
            cb_lighting_ceiling.Checked = true;
            cb_hvac_ceiling.Checked = true;
            cb_fa_floor.Checked = true;
            cb_fa_ceiling.Checked = true;
            cb_lc_ceiling.Checked = true;
            cb_coordination_ceiling.Checked = true;
            cb_coordination_floor.Checked = true;
        }

        private void UpdateCheckBoxAppearance(CheckBox checkBox)
        {
            if (checkBox == null)
                return;

            checkBox.ForeColor = checkBox.Checked ? Color.White : Color.Black;
        }

        private void ShowWarningMessage(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}