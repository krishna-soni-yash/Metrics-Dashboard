import genericService, { GenericService } from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ErrorMessages from '../common/ErrorMessages';
import { SubSiteListNames } from '../common/Constants';
import { IRAIDLogRepository } from './repositoryInterface/IRAIDLogRepository';
import { IRAIDLogs } from '../Models/IRAIDLogs';
//import { getListConfigurationBasedOnMetricLogs } from '../repositories/ObjectivesMasterRepository';
//import IObjectivesMasterRepository from './repositoriesInterface/IObjectivesMasterRepository';

/**
 * Repository for ProjectTypes list
 * Implements a simple cached fetch of Id/LinkTitle values
 */
export class RAIDLogRepository implements IRAIDLogRepository {
    private service: IGenericService;
    private cache: IRAIDLogs[] | null = null;
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




    public async getRiskValues(useCache: boolean = true, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<IRAIDLogs[]> {
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
                listTitle: SubSiteListNames.RAIDLog,
                //select: selectFields,
                pageSize: 5000,
                //orderBy: 'Created',
                //   $select=Id,Title,Created,Author/Title
                //   &$expand=Author
                //   &$orderby=Created desc
                //   &$top=1
                 // filter: "SelectType eq 'Risk'"
                // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',

            });

            const normalized = (items || []).map((it: any) => ({
               
                SelectType: it?.SelectType ?? '',
                RAIDId: it?.RAIDId??'',
                IdentificationDate: it?.IdentificationDate??'',
                RiskDescription: it?.RiskDescription??'',
                AssociatedGoal:it?.AssociatedGoal??'' ,
                RiskSource: it?.RiskSource??'',
                RiskCategory: it?.RiskCategory??'',
                Impact: it?.Impact??'',
                RiskPriority: it?.RiskPriority??'',
                ImpactValue: it?.ImpactValue,
                ProbabilityValue:it?.ProbabilityValue,
                RiskExposure: it?.RiskExposure

            })) as unknown as IRAIDLogs[];

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

const defaultInstance = new RAIDLogRepository();

export default defaultInstance;
export const MetricsRepo = defaultInstance;
export const getRiskValues = async (useCache: boolean = false, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<IRAIDLogs[]> => defaultInstance.getRiskValues(useCache, context, selectedStartDate, selectedEndDate);
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();