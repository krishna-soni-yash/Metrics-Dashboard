export interface IWorkLogManagement {
    Title: string;
    ReqTitle: string;
    ProjectType: string;
    WorkItemNo: string;
    Simple: Number;
    Medium: Number;
    Complex: Number;
    VeryComplex: Number;
    ComplexityPoints: Number;
    AppAndEnvAdjustmentFactor: Number;
    SkillAdjustmentFactor: Number;
    ReusabilityOfDesignAndCode: Number;
    ExtentOfAutomation: Number;
    AdjustedEffort: Number;
    BaseEffort: Number;
    CalculatedPlannedEffort: Number;
    ActualPlannedEffort: Number;
    PlannedStartDate: string
    PlannedEndDate: string;
    Status: string;
    Remarks: string;
    AdjustedComplexityPoint: Number;

}