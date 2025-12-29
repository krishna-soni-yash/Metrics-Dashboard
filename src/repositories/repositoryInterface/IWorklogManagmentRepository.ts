//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IGenericService from '../../services/IGenericServices';
import { IWorkLogManagement } from '../../Models/IWorkLogManagement';

export interface IWorklogManagmentRepository {
  
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
  getWorkLogManagementValues(useCache?: boolean, context?: WebPartContext, selectedProjectType?: string): Promise<IWorkLogManagement[]>;
}

export default IWorklogManagmentRepository;