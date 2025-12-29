//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IGenericService from '../../services/IGenericServices';
import {  IRAIDLogs } from '../../Models/IRAIDLogs';

export interface IRAIDLogRepository {
  
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
  getRiskValues(useCache?: boolean, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<IRAIDLogs[]>;}

export default IRAIDLogRepository;