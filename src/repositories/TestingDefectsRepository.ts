import genericService, { GenericService } from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ErrorMessages from '../common/ErrorMessages';
import { expandFields, selectedFields, SubSiteListNames } from '../common/Constants';
import { ITestingDefectsRepository } from './repositoryInterface/ITestingDefectsRepository';
import { ITestingDefects } from '../Models/ITestingDefects';
//import { getListConfigurationBasedOnMetricLogs } from '../repositories/ObjectivesMasterRepository';
//import IObjectivesMasterRepository from './repositoriesInterface/IObjectivesMasterRepository';

/**
 * Repository for ProjectTypes list
 * Implements a simple cached fetch of Id/LinkTitle values
 */
export class TestingDefectsRepository implements ITestingDefectsRepository {
    private service: IGenericService;
    private cache: ITestingDefects[] | null = null;
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




    public async getTestingDefectsValues(useCache: boolean = true, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<ITestingDefects[]> {
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
                listTitle: SubSiteListNames.TestingDefects,
                pageSize: 5000,
                select: selectedFields,
                expand: expandFields,
                //filter:'Identified_x0020_Phase eq \'Post Production\''


            });

            const normalized = (items || []).map((it: any) => ({
                Requirement: it?.Title,
                TestScenarioID: it?.Test_x0020_Scenario_x0020_ID,
                TestCaseID: it?.Test_x0020_Case_x0020_ID,
                DefectDescription: it?.Defect_x0020_Description,
                TestingType: it?.Review_x0020_method,
                DefectDetectedOn: it?.Defect_x0020_detected_x0020_on,
                DefectDetectedBy: it?.Defect_x0020_Detected_x0020_by.EMail,
                DefectStatus: it?.Defect_x0020_Status,
                DefectType: it?.Defect_x0020_Category,
                DefectClassification: it?.Defect_x0020_Classification_x002,
                DefectOriginPhase: it?.Injected_x0020_Phase,
                DefectDetectedPhase: it?.Identified_x0020_Phase,
                Severity: it?.Severity,
                Priority: it?.Priority,
                RootCause: it?.Root_x0020_Cause,
                DefectFixedBy: it?.Defect_x0020_Fixed_x0020_By,
                DefectClosureDate: it?.Defect_x0020_Closure_x0020_Date,
                Remarks: it?.Remarks



            })) as unknown as ITestingDefects[];

            this.cache = normalized;
            this.cacheTimestamp = now;

            return this.cache;
        } catch (error: any) {
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

const defaultInstance = new TestingDefectsRepository();

export default defaultInstance;
export const MetricsRepo = defaultInstance;
export const getTestingDefectsValues = async (useCache: boolean = false, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<ITestingDefects[]> => defaultInstance.getTestingDefectsValues(useCache, context, selectedStartDate, selectedEndDate);
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();