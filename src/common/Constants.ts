//import { IColumnConfig } from '../webparts/rootCauseAnalysis/components/RootCauseAnalysisTables/RCATable';

export default class ParentListNames {
  public static AppSettings: string = "AppSettings";
  public static ObjectivesMaster: string = "ObjectivesMaster";
  public static PPOApprovers: string = "PPOApprovers";
  public static ProjectType: string = "ProjectType";
  public static AssociatedPPM: string = "AssociatedPPM"
  public static Metrics: string = "Metrics";
  public static MetricsMailSender: string = "MetricsMailSender";
}

export class SubSiteListNames {
  public static ProjectMetrics: string = "ProjectMetrics";
  public static ProjectMetricLogs: string = "ProjectMetricLogs";
  public static RootCauseAnalysis: string = "Root Cause Analysis";
  public static WorkLogManagement: string = "WorkLogManagement";
  public static TaskManagement: string = "TaskManagement";
  public static ManagementEffortLog: string = "ManagementEffortLog";
  public static TaskManagementLog: string = "ManagementTaskLog";
  public static CSI: string = "Customer Satisfaction Index";
  public static RAIDLog: string = "RAIDLogs";
  public static TestingDefects: string = "Testing Defects";
  public static CodeReviewDefects: string = "Code Review Defects";
  public static AMSTicketEffortLog : string ="AMSTicketEffortLog"
}

//RCA Site Configuration

export class SiteConfiguration {
  public static readonly PARENT_LISTS = [
    ParentListNames.AppSettings,
    ParentListNames.ObjectivesMaster,
    ParentListNames.PPOApprovers,
    ParentListNames.ProjectType,
    ParentListNames.AssociatedPPM,
    ParentListNames.Metrics,
    ParentListNames.MetricsMailSender
  ];
}

export const selectedFields = [
  'Title',
  'Test_x0020_Scenario_x0020_ID',
  'Test_x0020_Case_x0020_ID',
  'Defect_x0020_Description',
  'Review_x0020_method',
  'Defect_x0020_detected_x0020_on',
  'Defect_x0020_Detected_x0020_by/EMail',
  'Defect_x0020_Status',
  'Defect_x0020_Category',
  'Defect_x0020_Classification_x002',
  'Injected_x0020_Phase',
  'Identified_x0020_Phase',
  'Severity',
  'Priority',
  'Root_x0020_Cause',
  'Defect_x0020_Fixed_x0020_By/EMail',
  'Defect_x0020_Closure_x0020_Date',
  'Remarks',
];
export const expandFields = [
'Defect_x0020_Detected_x0020_by',
'Defect_x0020_Fixed_x0020_By'

];
export const selectedFieldsCodeReview = [
  'Title',
    'Code_x0020_File_x002f__x0020_Cla',
    'Code_x0020_File_x002f__x0020_Cla0',
    'Code_x0020_File_x0020__x002d__x0/EMail',
    'Reviewer_x0020_Name/EMail',
    'Review_x0020_Iteration_x0020_Num',
    'Identified_x0020_Date',
    'Review_x0020_Completion_x0020_Da',
    'Defect_x0020_Description',
    'Code_x0020_Review_x0020_Checklis',
    'Review_x0020_Results',
    'Defect_x0020_Status',
    'Defect_x0020_Type',
    'Defect_x0020_Classification',
    'Defect_x0020_Origin_x0020_Phase',
    'Impacted_x0020_Components',
    'Correction_x0020__x002f__x0020_C',
    'Planned_x0020_Closure_x0020_Date',
    'Actual_x0020_Closure_x0020_Date',
    'Location_x0020_of_x0020_defect_x',
    'Severity',
    'Remarks'
];
export const expandFieldsCodeReview = [
  'Code_x0020_File_x0020__x002d__x0',
    'Reviewer_x0020_Name',
];

