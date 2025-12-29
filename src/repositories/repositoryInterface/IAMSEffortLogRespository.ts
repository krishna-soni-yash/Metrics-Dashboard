//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IGenericService from '../../services/IGenericServices';
import {  IAMSEffortLog } from '../../Models/IAMSEffortLog';

export interface IAMSEffortLogRepository {
  
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
  getIAMSEffortLogValues(useCache?: boolean, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<IAMSEffortLog[]>;}

export default IAMSEffortLogRepository;