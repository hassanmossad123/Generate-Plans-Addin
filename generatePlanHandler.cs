using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

public class GeneratePlanHandler : IExternalEventHandler
{
    // ========== CONSTANT DEFINITIONS ==========
    public const string POWER_SUB_DISCIPLINE = "Power";
    public const string LIGHTING_SUB_DISCIPLINE = "Lighting";
    public const string HVAC_SUB_DISCIPLINE = "HVAC";
    public const string FIRE_ALARM_SUB_DISCIPLINE = "FA";
    public const string LIGHT_CURRENT_SUB_DISCIPLINE = "LC";
    public const string COORDINATION_SUB_DISCIPLINE = "Coordination";

    private const string FLOOR_PLAN_SUFFIX = "Floor";
    private const string CEILING_PLAN_SUFFIX = "Ceiling";

    // View Range Constants (in mm)
    private const double CUT_PLANE_OFFSET = 1200;
    private const double VIEW_DEPTH_OFFSET = 0;

    // ========== PRIVATE FIELDS ==========
    private Document _document;
    private Level _selectedLevel;
    private Dictionary<string, (bool CreateFloorPlan, bool CreateCeilingPlan)> _disciplineSettings;

    // ========== CONSTRUCTOR ==========
    public GeneratePlanHandler(Document document)
    {
        _document = document;
    }

    // ========== PUBLIC METHODS ==========
    public void Execute(UIApplication application)
    {
        try
        {
            using (TransactionGroup transactionGroup = new TransactionGroup(_document, "Generate Plans"))
            {
                transactionGroup.Start();
                List<string> createdPlans = GenerateSelectedPlans();
                transactionGroup.Assimilate();
                DisplayGenerationResults(createdPlans);
            }
        }
        catch (Exception exception)
        {
            TaskDialog.Show("Generation Error", $"Failed to generate plans: {exception.Message}");
        }
    }

    public void Initialize(Level level, Dictionary<string, (bool CreateFloorPlan, bool CreateCeilingPlan)> disciplineSettings)
    {
        _selectedLevel = level;
        _disciplineSettings = disciplineSettings ?? new Dictionary<string, (bool, bool)>();
    }

    public string GetName()
    {
        return "Generate Floor/Ceiling Plans";
    }

    // ========== PRIVATE CORE METHODS ==========
    private List<string> GenerateSelectedPlans()
    {
        List<string> createdPlanNames = new List<string>();
        List<string> skippedDuplicatePlans = new List<string>();
        List<Level> targetLevels = GetTargetLevels();

        foreach (Level targetLevel in targetLevels)
        {
            foreach (KeyValuePair<string, (bool CreateFloorPlan, bool CreateCeilingPlan)> disciplineSetting in _disciplineSettings)
            {
                string discipline = disciplineSetting.Key;

                if (disciplineSetting.Value.CreateFloorPlan)
                {
                    var result = CreatePlanView(targetLevel, ViewFamily.FloorPlan, FLOOR_PLAN_SUFFIX, discipline);
                    ProcessCreationResult(result, createdPlanNames, skippedDuplicatePlans);
                }

                if (disciplineSetting.Value.CreateCeilingPlan)
                {
                    var result = CreatePlanView(targetLevel, ViewFamily.CeilingPlan, CEILING_PLAN_SUFFIX, discipline);
                    ProcessCreationResult(result, createdPlanNames, skippedDuplicatePlans);
                }
            }
        }

        NotifyAboutSkippedPlans(skippedDuplicatePlans);
        return createdPlanNames;
    }

    private (bool Success, string PlanName, string ErrorMessage) CreatePlanView(Level level, ViewFamily viewFamily, string planType, string discipline)
    {
        string planName = $"{planType} {discipline} - {level.Name}";

        if (PlanViewExists(planName))
        {
            return (false, planName, "View already exists");
        }

        try
        {
            using (Transaction creationTransaction = new Transaction(_document, $"Create {planType} Plan"))
            {
                creationTransaction.Start();

                ViewFamilyType viewFamilyType = GetViewFamilyType(viewFamily);
                ViewPlan newPlanView = ViewPlan.Create(_document, viewFamilyType.Id, level.Id);
                newPlanView.Name = planName;

                ConfigurePlanViewRange(newPlanView);
                ApplyViewTemplate(newPlanView, $"{planType} {discipline} Template", discipline);

                creationTransaction.Commit();
                return (true, planName, null);
            }
        }
        catch (Exception exception)
        {
            return (false, planName, exception.Message);
        }
    }

    // ========== VIEW TEMPLATE METHODS ==========
    private void ApplyViewTemplate(ViewPlan planView, string templateName, string discipline)
    {
        View existingTemplate = new FilteredElementCollector(_document)
            .OfClass(typeof(View))
            .Cast<View>()
            .FirstOrDefault(view => view.IsTemplate && view.Name == templateName);

        if (existingTemplate == null)
        {
            existingTemplate = planView.CreateViewTemplate();
            existingTemplate.Name = templateName;
        }

        ConfigureDisciplineTemplate(existingTemplate, discipline);
        planView.ViewTemplateId = existingTemplate.Id;
        AssignSubDisciplineParameter(existingTemplate, discipline);
    }

    private void ConfigureDisciplineTemplate(View template, string discipline)
    {
        Document templateDocument = template.Document;
        Categories documentCategories = templateDocument.Settings.Categories;

        switch (discipline)
        {
            case POWER_SUB_DISCIPLINE:
                ConfigurePowerTemplate(template, documentCategories);
                break;
            case LIGHTING_SUB_DISCIPLINE:
                ConfigureLightingTemplate(template, documentCategories);
                break;
            case HVAC_SUB_DISCIPLINE:
                ConfigureHVACTemplate(template, documentCategories);
                break;
            case LIGHT_CURRENT_SUB_DISCIPLINE:
                ConfigureLightCurrentTemplate(template, documentCategories);
                break;
            case FIRE_ALARM_SUB_DISCIPLINE:
                ConfigureFireAlarmTemplate(template, documentCategories);
                break;
        }

        template.DetailLevel = ViewDetailLevel.Fine;
        template.Discipline = ViewDiscipline.Electrical;
    }

    private void AssignSubDisciplineParameter(View template, string discipline)
    {
        foreach (Parameter parameter in template.Parameters)
        {
            if (parameter.Definition.Name == "Sub-Discipline")
            {
                parameter.Set(discipline);
                break;
            }
        }
    }

    // ========== VIEW RANGE CONFIGURATION ==========
    private void ConfigurePlanViewRange(ViewPlan planView)
    {
        PlanViewRange viewRange = planView.GetViewRange();

        Level upperLevel = new FilteredElementCollector(_document)
            .OfClass(typeof(Level))
            .Cast<Level>()
            .Where(level => level.Elevation > planView.GenLevel.Elevation)
            .OrderBy(level => level.Elevation)
            .FirstOrDefault();

        if (upperLevel != null)
        {
            viewRange.SetLevelId(PlanViewPlane.TopClipPlane, upperLevel.Id);
            viewRange.SetOffset(PlanViewPlane.TopClipPlane, VIEW_DEPTH_OFFSET);
        }

        planView.SetViewRange(viewRange);
    }

    // ========== DISCIPLINE-SPECIFIC CONFIGURATIONS ==========
    private void ConfigurePowerTemplate(View template, Categories categories)
    {
        List<BuiltInCategory> hiddenCategories = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_CommunicationDevices,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_CableTrayFitting,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ConduitFitting,
            BuiltInCategory.OST_DataDevices,
            BuiltInCategory.OST_FireAlarmDevices,
            BuiltInCategory.OST_FireProtection,
            BuiltInCategory.OST_LightingDevices,
            BuiltInCategory.OST_LightingFixtures,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_MedicalEquipment,
            BuiltInCategory.OST_NurseCallDevices,
            BuiltInCategory.OST_Parts,
            BuiltInCategory.OST_SecurityDevices,
            BuiltInCategory.OST_TelephoneDevices,
            
            // Tag Categories
            BuiltInCategory.OST_CommunicationDeviceTags,
            BuiltInCategory.OST_CableTrayTags,
            BuiltInCategory.OST_ConduitTags,
            BuiltInCategory.OST_DataDeviceTags,
            BuiltInCategory.OST_FireAlarmDeviceTags,
            BuiltInCategory.OST_FireProtectionTags,
            BuiltInCategory.OST_LightingDeviceTags,
            BuiltInCategory.OST_LightingFixtureTags,
            BuiltInCategory.OST_MechanicalEquipmentTags,
            BuiltInCategory.OST_MedicalEquipmentTags,
            BuiltInCategory.OST_NurseCallDeviceTags,
            BuiltInCategory.OST_SecurityDeviceTags,
            BuiltInCategory.OST_TelephoneDeviceTags,
        };

        ApplyCategoryVisibilitySettings(template, categories, hiddenCategories);
    }

    private void ConfigureLightingTemplate(View template, Categories categories)
    {
        List<BuiltInCategory> hiddenCategories = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_CommunicationDevices,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_DataDevices,
            BuiltInCategory.OST_ElectricalFixtures,
            BuiltInCategory.OST_FireAlarmDevices,
            BuiltInCategory.OST_FireProtection,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_MedicalEquipment,
            BuiltInCategory.OST_NurseCallDevices,
            BuiltInCategory.OST_Parts,
            BuiltInCategory.OST_SecurityDevices,
            BuiltInCategory.OST_TelephoneDevices,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ConduitFitting,
            
            // Tag Categories
            BuiltInCategory.OST_CommunicationDeviceTags,
            BuiltInCategory.OST_CableTrayTags,
            BuiltInCategory.OST_DataDeviceTags,
            BuiltInCategory.OST_ElectricalFixtureTags,
            BuiltInCategory.OST_FireAlarmDeviceTags,
            BuiltInCategory.OST_FireProtectionTags,
            BuiltInCategory.OST_MechanicalEquipmentTags,
            BuiltInCategory.OST_MedicalEquipmentTags,
            BuiltInCategory.OST_NurseCallDeviceTags,
            BuiltInCategory.OST_SecurityDeviceTags,
            BuiltInCategory.OST_TelephoneDeviceTags,
            BuiltInCategory.OST_ConduitTags,
            BuiltInCategory.OST_ConduitFittingTags,
        };

        ApplyCategoryVisibilitySettings(template, categories, hiddenCategories);
    }

    private void ConfigureHVACTemplate(View template, Categories categories)
    {
        List<BuiltInCategory> hiddenCategories = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_LightingFixtures,
            BuiltInCategory.OST_LightingDevices,
            BuiltInCategory.OST_FireAlarmDevices,
            BuiltInCategory.OST_CommunicationDevices,
            BuiltInCategory.OST_DataDevices,
            BuiltInCategory.OST_TelephoneDevices,
            BuiltInCategory.OST_SecurityDevices,
            BuiltInCategory.OST_NurseCallDevices,
            BuiltInCategory.OST_MedicalEquipment,
            BuiltInCategory.OST_FireProtection,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_CableTrayFitting,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ConduitFitting,
            BuiltInCategory.OST_Parts,
            
            // Tag Categories
            BuiltInCategory.OST_LightingFixtureTags,
            BuiltInCategory.OST_LightingDeviceTags,
            BuiltInCategory.OST_FireAlarmDeviceTags,
            BuiltInCategory.OST_CommunicationDeviceTags,
            BuiltInCategory.OST_DataDeviceTags,
            BuiltInCategory.OST_TelephoneDeviceTags,
            BuiltInCategory.OST_SecurityDeviceTags,
            BuiltInCategory.OST_NurseCallDeviceTags,
            BuiltInCategory.OST_MedicalEquipmentTags,
            BuiltInCategory.OST_FireProtectionTags,
            BuiltInCategory.OST_CableTrayTags,
            BuiltInCategory.OST_ConduitTags,
        };

        ApplyCategoryVisibilitySettings(template, categories, hiddenCategories);
        template.Discipline = ViewDiscipline.Mechanical;
    }

    private void ConfigureLightCurrentTemplate(View template, Categories categories)
    {
        List<BuiltInCategory> hiddenCategories = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_ElectricalFixtures,
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_LightingFixtures,
            BuiltInCategory.OST_LightingDevices,
            BuiltInCategory.OST_FireAlarmDevices,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_CableTrayFitting,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ConduitFitting,
            BuiltInCategory.OST_MedicalEquipment,
            
            // Tag Categories
            BuiltInCategory.OST_ElectricalFixtureTags,
            BuiltInCategory.OST_ElectricalEquipmentTags,
            BuiltInCategory.OST_LightingFixtureTags,
            BuiltInCategory.OST_LightingDeviceTags,
            BuiltInCategory.OST_CableTrayTags,
            BuiltInCategory.OST_ConduitTags,
            BuiltInCategory.OST_FireAlarmDeviceTags,
            BuiltInCategory.OST_MedicalEquipmentTags,
        };

        ApplyCategoryVisibilitySettings(template, categories, hiddenCategories);
    }

    private void ConfigureFireAlarmTemplate(View template, Categories categories)
    {
        List<BuiltInCategory> hiddenCategories = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_ElectricalFixtures,
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_LightingFixtures,
            BuiltInCategory.OST_LightingDevices,
            BuiltInCategory.OST_CommunicationDevices,
            BuiltInCategory.OST_DataDevices,
            BuiltInCategory.OST_TelephoneDevices,
            BuiltInCategory.OST_SecurityDevices,
            BuiltInCategory.OST_NurseCallDevices,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_CableTrayFitting,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ConduitFitting,
            BuiltInCategory.OST_MedicalEquipment,
            BuiltInCategory.OST_Parts,
            
            // Tag Categories
            BuiltInCategory.OST_ElectricalFixtureTags,
            BuiltInCategory.OST_ElectricalEquipmentTags,
            BuiltInCategory.OST_LightingFixtureTags,
            BuiltInCategory.OST_LightingDeviceTags,
            BuiltInCategory.OST_CommunicationDeviceTags,
            BuiltInCategory.OST_DataDeviceTags,
            BuiltInCategory.OST_TelephoneDeviceTags,
            BuiltInCategory.OST_SecurityDeviceTags,
            BuiltInCategory.OST_NurseCallDeviceTags,
            BuiltInCategory.OST_CableTrayTags,
            BuiltInCategory.OST_ConduitTags,
            BuiltInCategory.OST_MedicalEquipmentTags,
        };

        ApplyCategoryVisibilitySettings(template, categories, hiddenCategories);
    }

    private void ApplyCategoryVisibilitySettings(View template, Categories categories, List<BuiltInCategory> hiddenCategories)
    {
        foreach (BuiltInCategory category in hiddenCategories)
        {
            HideCategory(template, categories, category);
        }
    }

    private void HideCategory(View template, Categories categories, BuiltInCategory category)
    {
        try
        {
            Category revitCategory = categories.get_Item(category);
            if (revitCategory != null)
            {
                template.SetCategoryHidden(revitCategory.Id, true);
            }
        }
        catch
        {
            // Category doesn't exist in this project, skip it
        }
    }

    // ========== HELPER METHODS ==========
    private List<Level> GetTargetLevels()
    {
        if (_selectedLevel == null)
        {
            return GetAllDocumentLevels().OrderBy(level => level.Elevation).ToList();
        }
        return new List<Level> { _selectedLevel };
    }

    private void ProcessCreationResult((bool Success, string PlanName, string ErrorMessage) result,
        List<string> createdPlans, List<string> skippedPlans)
    {
        if (result.Success)
        {
            createdPlans.Add(result.PlanName);
        }
        else if (result.ErrorMessage == "View already exists")
        {
            skippedPlans.Add(result.PlanName);
        }
    }

    private void NotifyAboutSkippedPlans(List<string> skippedPlans)
    {
        if (skippedPlans.Count == 0) return;

        StringBuilder notificationMessage = new StringBuilder();
        notificationMessage.AppendLine($"⚠️ {skippedPlans.Count} plan(s) already exist and were skipped:");
        notificationMessage.AppendLine();

        foreach (string planName in skippedPlans)
        {
            notificationMessage.AppendLine($"• {planName}");
        }

        notificationMessage.AppendLine();
        notificationMessage.AppendLine("These plans were not created to avoid duplicates.");

        TaskDialog.Show("Duplicate Plans Skipped", notificationMessage.ToString());
    }

    private bool PlanViewExists(string planName)
    {
        return new FilteredElementCollector(_document)
            .OfClass(typeof(View))
            .Cast<View>()
            .Any(view => view.Name.Equals(planName, StringComparison.OrdinalIgnoreCase));
    }

    private ViewFamilyType GetViewFamilyType(ViewFamily viewFamily)
    {
        return new FilteredElementCollector(_document)
            .OfClass(typeof(ViewFamilyType))
            .Cast<ViewFamilyType>()
            .FirstOrDefault(viewFamilyType => viewFamilyType.ViewFamily == viewFamily);
    }

    public List<Level> GetAllDocumentLevels()
    {
        return new FilteredElementCollector(_document)
            .OfClass(typeof(Level))
            .Cast<Level>()
            .ToList();
    }

    private void DisplayGenerationResults(List<string> createdPlans)
    {
        if (createdPlans.Count == 0)
        {
            TaskDialog.Show("Generation Complete", "No new plans were created.");
            return;
        }

        StringBuilder resultsMessage = new StringBuilder();
        resultsMessage.AppendLine($"✅ Successfully created {createdPlans.Count} plan(s):");
        resultsMessage.AppendLine();

        foreach (string planName in createdPlans)
        {
            resultsMessage.AppendLine($"• {planName}");
        }

        TaskDialog.Show("Generation Successful", resultsMessage.ToString());
    }
}