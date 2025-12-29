//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IGenericService from '../../services/IGenericServices';
import {  IManagementEffortLog } from '../../Models/IManagementEffortLog';

export interface IManagemnetEffortLogRepository {
  
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
  getManagementEffortLogValues(useCache?: boolean, context?: WebPartContext, selectedProjectType?: string): Promise<IManagementEffortLog[]>;
}

export default IManagemnetEffortLogRepository;