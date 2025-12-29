import genericService, { GenericService } from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ErrorMessages from '../common/ErrorMessages';
import { SubSiteListNames } from '../common/Constants';
import { IWorklogManagmentRepository } from './repositoryInterface/IWorklogManagmentRepository';
import { IWorkLogManagement } from '../Models/IWorkLogManagement';
//import { getListConfigurationBasedOnMetricLogs } from '../repositories/ObjectivesMasterRepository';
//import IObjectivesMasterRepository from './repositoriesInterface/IObjectivesMasterRepository';

/**
 * Repository for ProjectTypes list
 * Implements a simple cached fetch of Id/LinkTitle values
 */
export class WorklogManagmentRepository implements IWorklogManagmentRepository {
    private service: IGenericService;
    private cache: IWorkLogManagement[] | null = null;
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




    public async getWorkLogManagementValues(useCache: boolean = true, context?: WebPartContext, selectedProjectType?: string): Promise<IWorkLogManagement[]> {
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
                listTitle: SubSiteListNames.WorkLogManagement,
                //select: selectFields,
                pageSize: 2000,
                filter: 'ProjectType eq (\'' + (selectedProjectType) + '\')',

                // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',

            });

            const normalized = (items || []).map((it: any) => ({
                Title: it?.Title ?? '',
                ReqTitle: it?.ReqTitle ?? '',
                ProjectType: it?.ProjectType ?? '',
                WorkItemNo: it?.WorkItemNo ?? '',
                Simple: it?.Simple ?? 0,
                Medium: it?.Medium ?? 0,
                Complex: it?.Complex ?? 0,
                VeryComplex: it?.VeryComplex ?? 0,
                ComplexityPoints: it?.ComplexityPoints ?? 0,
                AppAndEnvAdjustmentFactor: it?.AppAndEnvAdjustmentFactor ?? 0,
                SkillAdjustmentFactor: it?.SkillAdjustmentFactor ?? 0,
                ReusabilityOfDesignAndCode: it?.ReusabilityOfDesignAndCode ?? 0,
                ExtentOfAutomation: it?.ExtentOfAutomation ?? 0,
                AdjustedEffort: it?.AdjustedEffort ?? 0,
                BaseEffort: it?.BaseEffort ?? 0,
                CalculatedPlannedEffort: it?.CalculatedPlannedEffort ?? 0,
                ActualPlannedEffort: it?.ActualPlannedEffort ?? 0,
                PlannedStartDate: it?.PlannedStartDate ?? '',
                PlannedEndDate: it?.PlannedEndDate ?? '',
                Status: it?.Status ?? '',
                Remarks: it?.Remarks ?? '',
                AdjustedComplexityPoint: it?.AdjustedComplexityPoint ?? 0,

            })) as unknown as IWorkLogManagement[];

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

const defaultInstance = new WorklogManagmentRepository();

export default defaultInstance;
export const MetricsRepo = defaultInstance;
export const getWorkLogManagementValues = async (useCache: boolean = false, context?: WebPartContext, selectedProjectType?: string): Promise<IWorkLogManagement[]> => defaultInstance.getWorkLogManagementValues(useCache, context, selectedProjectType);
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();