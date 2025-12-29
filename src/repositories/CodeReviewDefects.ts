import genericService, { GenericService } from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ErrorMessages from '../common/ErrorMessages';
import { expandFields, expandFieldsCodeReview, selectedFields, selectedFieldsCodeReview, SubSiteListNames } from '../common/Constants';
import { ICodeReviewDefectsRepository } from './repositoryInterface/ICodeReviewDefectsRepository';
import { ICodeReviewDefects } from '../Models/ICodeReviewDefects';
//import { getListConfigurationBasedOnMetricLogs } from '../repositories/ObjectivesMasterRepository';
//import IObjectivesMasterRepository from './repositoriesInterface/IObjectivesMasterRepository';

/**
 * Repository for ProjectTypes list
 * Implements a simple cached fetch of Id/LinkTitle values
 */
export class CodeReviewDefectRepository implements ICodeReviewDefectsRepository {
    private service: IGenericService;
    private cache: ICodeReviewDefects[] | null = null;
    private cacheTimestamp = 0;
    private readonly CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

    constructor(service?: IGenericService) {
        this.service = service ?? genericService;

    }

    public setService(service: IGenericService): void {
        this.service = service;
    }
    // private normalizeSiteUrl(value?: string): string {
    //     return (value || '').trim().replace(/\/+$/, '').toLowerCase();
    //   }

    // private ActiveVersion: IMetrics[] | null = null;
    //private VersionId: number | undefined = undefined;




    public async getCodeReviewDefectsValues(useCache: boolean = true, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<ICodeReviewDefects[]> {
        const now = Date.now();

        if (useCache && this.cache && (now - this.cacheTimestamp) < this.CACHE_DURATION) {
            return this.cache;
        }

        if (!context) {
            throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
        }

        try {
            const genericServiceInstance: IGenericService = new GenericService(undefined, context);
            genericServiceInstance.init(undefined, context);
            // const selectFields: string[] = ['Id', 'LinkTitle','ProjectType','IsActive'];
            //const listConfig = await getListConfigurationBasedOnMetricLogs(context);

            //this.ActiveVersion = await this.getApprovedProjectlogs(true, context);
            // if (this.ActiveVersion && this.ActiveVersion.length > 0) {
            //     this.VersionId = this.ActiveVersion[0].ID;
            // }
            const items = await this.service.fetchAllItems<any>({
                context,
                listTitle: SubSiteListNames.CodeReviewDefects,
                pageSize: 5000,
                select: selectedFieldsCodeReview,
                expand: expandFieldsCodeReview,
                //filter:'Identified_x0020_Phase eq \'Post Production\''


            });

            const normalized = (items || []).map((it: any) => ({
                RequirementID: it?.Title,
                CodeFileClassName: it?.Code_x0020_File_x002f__x0020_Cla,
                CodeFileClassSize: it?.Code_x0020_File_x002f__x0020_Cla0,
                CodeFileAuthor: it?.Code_x0020_File_x0020__x002d__x0,
                ReviewerName: it?.Reviewer_x0020_Name,
                ReviewIterationNumber: it?.Review_x0020_Iteration_x0020_Num,
                IdentifiedDate: it?.Identified_x0020_Date,
                ReviewCompletionDate: it?.Review_x0020_Completion_x0020_Da,
                //Defect_x0020_Description
                DefectDescription: it?.Defect_x0020_Description,
                //Code_x0020_Review_x0020_Checklis
                CodeReviewChecklist: it?.Code_x0020_Review_x0020_Checklis,
                //Review_x0020_Results
                ReviewResults: it?.Review_x0020_Results,
                //Defect_x0020_Status
                DefectStatus: it?.Defect_x0020_Status,
                //Defect_x0020_Type
                DefectType: it?.Defect_x0020_Type,
                //Defect_x0020_Classification
                DefectClassification: it?.Defect_x0020_Classification,
                //Defect_x0020_Origin_x0020_Phase
                DefectOriginPhase: it?.Defect_x0020_Origin_x0020_Phase,
                //Impacted Components
                impactedComponents: it?.Impacted_x0020_Components,
                //Correction_x0020__x002f__x0020_C
                CorrectionCorrectiveAction: it?.Correction_x0020__x002f__x0020_C,
                //Planned_x0020_Closure_x0020_Date
                PlannedClosureDate: it?.Planned_x0020_Closure_x0020_Date,
                //Actual_x0020_Closure_x0020_Date
                ActualClosureDate: it?.Actual_x0020_Closure_x0020_Date,
                //Location_x0020_of_x0020_defect_x
                LocationOfDefectInCode: it?.Location_x0020_of_x0020_defect_x,
                Severity: it?.Severity,
                Remarks: it?.Remarks

            })) as unknown as ICodeReviewDefects[];

        this.cache = normalized;
        this.cacheTimestamp = now;

        return this.cache;
    } catch(error: any) {
        throw new Error('Failed to fetch ProjectType: ' + (error?.message || error));
    }
}



    public refresh(): void {
    this.cache = null;
    this.cacheTimestamp = 0;
}

    public getCacheStatus(): { cached: boolean; itemCount: number; age: number } {
    const now = Date.now();
    return {
        cached: this.cache !== null,
        itemCount: this.cache?.length || 0,
        age: this.cache ? now - this.cacheTimestamp : 0
    };
}
}

const defaultInstance = new CodeReviewDefectRepository();

export default defaultInstance;
export const MetricsRepo = defaultInstance;
export const getCodeReviewDefectsValues = async (useCache: boolean = false, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<ICodeReviewDefects[]> => defaultInstance.getCodeReviewDefectsValues(useCache, context, selectedStartDate, selectedEndDate);
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();