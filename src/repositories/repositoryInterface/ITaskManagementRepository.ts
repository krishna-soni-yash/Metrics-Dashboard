//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IGenericService from '../../services/IGenericServices';
import {  ITaskManagement } from '../../Models/ITaskManagement';

export interface ITaskManagementRepository {
  
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
  getTaskManagementValues(useCache?: boolean, context?: WebPartContext, selectedProjectType?: string): Promise<ITaskManagement[]>;
}

export default ITaskManagementRepository;