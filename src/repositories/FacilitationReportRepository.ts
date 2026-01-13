import genericService, { GenericService } from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ErrorMessages from '../common/ErrorMessages';
import { SubSiteListNames } from '../common/Constants';
import { IFacilitationReportRepository } from './repositoryInterface/IFacilitationReportInterface';
import { IFacilitationReport } from '../Models/IFacilationReport';
//import { getListConfigurationBasedOnMetricLogs } from '../repositories/ObjectivesMasterRepository';
//import IObjectivesMasterRepository from './repositoriesInterface/IObjectivesMasterRepository';

/**
 * Repository for ProjectTypes list
 * Implements a simple cached fetch of Id/LinkTitle values
 */
export class FacilitationReportRepository implements IFacilitationReportRepository {
    private service: IGenericService;
    private cache: IFacilitationReport[] | null = null;
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




    public async getFacilitationValues(useCache: boolean = true, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<IFacilitationReport[]> {
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
                listTitle: SubSiteListNames.FacilitationReport,
                //select: selectFields,
                pageSize: 5000,
                filter:"Status eq 'open'",
               // orderBy: 'Created',
                //   $select=Id,Title,Created,Author/Title
                //   &$expand=Author
                //   &$orderby=Created desc
                //   &$top=1
                //  filter: "CSATAquiredDate ge datetime'" + selectedStartDate + "' and (CSATAquiredDate le datetime'" + selectedEndDate + "')",
                // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',

            });

            const normalized = (items || []).map((it: any) => ({
                ID: it?.ID,
                Category: it?.Category,
                Finding: it?.Finding,
                FindingDate: it?.FindingDate,
                ClosureDate: it?.ClosureDate,
                Status: it?.Status

            })) as unknown as IFacilitationReport[];

            this.cache = normalized;
            this.cacheTimestamp = now;

            return this.cache;
        } catch (error: any) {
            throw new Error('Failed to fetch FacilitationReport: ' + (error?.message || error));
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

const defaultInstance = new FacilitationReportRepository();

export default defaultInstance;
export const MetricsRepo = defaultInstance;
export const getFacilitationValues = async (useCache: boolean = false, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<IFacilitationReport[]> => defaultInstance.getFacilitationValues(useCache, context, selectedStartDate, selectedEndDate);
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();