import genericService, { GenericService } from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ErrorMessages from '../common/ErrorMessages';
import { SubSiteListNames } from '../common/Constants';
import { ITaskManagementRepository } from './repositoryInterface/ITaskManagementRepository';
import { ITaskManagement } from '../Models/ITaskManagement';
//import { getListConfigurationBasedOnMetricLogs } from '../repositories/ObjectivesMasterRepository';
//import IObjectivesMasterRepository from './repositoriesInterface/IObjectivesMasterRepository';

/**
 * Repository for ProjectTypes list
 * Implements a simple cached fetch of Id/LinkTitle values
 */
export class TaskManagementRepository implements ITaskManagementRepository {
    private service: IGenericService;
    private cache: ITaskManagement[] | null = null;
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




    public async getTaskManagementValues(useCache: boolean = true, context?: WebPartContext, selectedProjectType?: string): Promise<ITaskManagement[]> {
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
                listTitle: SubSiteListNames.TaskManagement,
                //select: selectFields,
                pageSize: 2000,
               // filter: 'ProjectType eq (\'' + (selectedProjectType) + '\')',

                // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',
                expand: ['WorkItemNo', 'AssignedTo'],
                select: ['WorkItemNo/WorkItemNo', 'ActualEffort', 'ActualEndDate', 'ActualStartDate',  'AssignedTo/Id', 'AssignedTo/Title','AssignedTo/EMail', 'Remarks', 'TaskStatus', 'TaskType','LinkTitle']

            });

            const normalized = (items || []).map((it: any) => ({
                WorkItemNo: it?.WorkItemNo.WorkItemNo ?? '',
               // Title:it?.LinkTitle??'',
                ReqTitle: it?.LinkTitle ?? '',
                ActualEffort: it?.ActualEffort ?? 0,
                ActualEndDate: it?.ActualEndDate ?? '',
                ActualStartDate: it?.ActualStartDate ?? '',
                AssignedTo: it?.AssignedTo.EMail ?? '',
                Remarks: it?.Remarks ?? '',
                TaskStatus: it?.TaskStatus ?? '',
                TaskType: it?.TaskType ?? '',

            })) as unknown as ITaskManagement[];

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

const defaultInstance = new TaskManagementRepository();

export default defaultInstance;
export const MetricsRepo = defaultInstance;
export const getTaskManagementValues = async (useCache: boolean = false, context?: WebPartContext, selectedProjectType?: string): Promise<ITaskManagement[]> => defaultInstance.getTaskManagementValues(useCache, context, selectedProjectType);
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();