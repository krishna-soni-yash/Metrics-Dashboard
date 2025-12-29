export interface ICodeReviewDefects {
    // id: number;
    RequirementID: string;
    //Code_x0020_File_x002f__x0020_Cla
    CodeFileClassName: string;
    //Code_x0020_File_x002f__x0020_Cla0
    CodeFileClassSize: string;
    //Code_x0020_File_x0020__x002d__x0
    CodeFileAuthor: string;
    //Reviewer_x0020_Name
    ReviewerName: string;
    //Review_x0020_Iteration_x0020_Num
    ReviewIterationNumber: string;
    //Identified_x0020_Date
    IdentifiedDate: string;
    //Review_x0020_Completion_x0020_Da
    ReviewCompletionDate: string;
    //Defect_x0020_Description
    DefectDescription: string;
    //Code_x0020_Review_x0020_Checklis
    CodeReviewChecklist: string;
    //Review_x0020_Results
    ReviewResults: string;

    DefectDetectedOn: string;
    DefectDetectedBy: string;
    //Defect_x0020_Status
    DefectStatus: string;
    //Defect_x0020_Type
    DefectType: string;
    //Defect_x0020_Classification
    DefectClassification: string;
    //Defect_x0020_Origin_x0020_Phase
    DefectOriginPhase: string;
    //Impacted Components
    impactedComponents: string;
    //Correction_x0020__x002f__x0020_C
    CorrectionCorrectiveAction: string;
    //Planned_x0020_Closure_x0020_Date
    PlannedClosureDate: string;
    //Actual_x0020_Closure_x0020_Date
    ActualClosureDate: string;
    //Location_x0020_of_x0020_defect_x
    LocationOfDefectInCode: string;
    Severity: string;
    Remarks: string;

}