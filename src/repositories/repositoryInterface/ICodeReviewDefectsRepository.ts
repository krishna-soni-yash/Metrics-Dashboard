//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IGenericService from '../../services/IGenericServices';
import {  ICodeReviewDefects } from '../../Models/ICodeReviewDefects';

export interface ICodeReviewDefectsRepository {
  
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
  getCodeReviewDefectsValues(useCache?: boolean, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<ICodeReviewDefects[]>;}

export default ICodeReviewDefectsRepository;