export interface ITestingDefects {
    // id: number;
    Requirement: string;
    TestScenarioID:string;
    TestCaseID:string;
    DefectDescription: string;
    TestingType: string;
    DefectDetectedOn: string;
    DefectDetectedBy: string;
    DefectStatus: string;
    //DefectType: string;
   // DefectClassification: string;
    DefectOriginPhase: string;
    DefectDetectedPhase: string;
    Severity: string;
    Priority: string;
    RootCause: string;
    DefectFixedBy: string;
    DefectClosureDate: string;
    Remarks: string;

}