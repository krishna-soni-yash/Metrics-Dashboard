//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IGenericService from '../../services/IGenericServices';
import { IManagementTaskLog } from '../../Models/IManagementTaskLog';

export interface ITaskManagementLogRepository {
  
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
  getTaskManagementLogValues(useCache?: boolean, context?: WebPartContext, selectedProjectType?: string): Promise<IManagementTaskLog[]>;
}

export default ITaskManagementLogRepository;