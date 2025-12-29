import GenericService from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPPOApprovers } from '../Models/IPPOApprovers';
import ParentListNames from '../common/Constants';
import ErrorMessages from '../common/ErrorMessages';
import IPPOApproversRepository from '../repositories/repositoryInterface/IPPOApproversRepository';

/**
 * Repository for PPOApprovers list
 */
export class PPOApproversRepository implements IPPOApproversRepository {
  private service: IGenericService;
  private cache: IPPOApprovers[] | null = null;
  private cacheTimestamp = 0;
  private cacheKey: string | null = null;
  private readonly CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

  constructor(service?: IGenericService) {
    this.service = service ?? GenericService;
  }

  public setService(service: IGenericService): void {
    this.service = service;
  }

  private normalizeSiteUrl(value?: string): string {
    return (value || '').trim().replace(/\/+$/, '').toLowerCase();
  }

  private normalizeEmail(value?: string): string {
    return (value || '').trim().toLowerCase();
  }

  private escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }

  private deriveSiteUrlCandidates(siteUrl: string): string[] {
    const trimmed = (siteUrl || '').trim();
    if (!trimmed) {
      return [];
    }

    const withoutTrailingSlash = trimmed.replace(/\/+$/, '');
    const candidates: Record<string, boolean> = {};

    const addCandidate = (value: string): void => {
      if (!value) {
        return;
      }
      candidates[value] = true;
    };

    addCandidate(trimmed);
    if (withoutTrailingSlash) {
      addCandidate(withoutTrailingSlash);
      addCandidate(`${withoutTrailingSlash}/`);
    }

    return Object.keys(candidates);
  }

  public async getAllItems(useCache: boolean = true, context?: WebPartContext): Promise<IPPOApprovers[]> {
    const now = Date.now();

    if (!context) {
      throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_PPOAPPROVERS || 'Web part context required');
    }

    const currentSiteUrlRaw = context.pageContext?.web?.absoluteUrl ?? '';
    const normalizedSiteUrl = this.normalizeSiteUrl(currentSiteUrlRaw);
    const currentUserEmailNormalized = this.normalizeEmail(context.pageContext?.user?.email);
    const cacheKey = `${normalizedSiteUrl || 'all'}::${currentUserEmailNormalized || 'anonymous'}`;

    if (useCache && this.cache && this.cacheKey === cacheKey && (now - this.cacheTimestamp) < this.CACHE_DURATION) {
      return this.cache;
    }

    try {
      const selectFields = [
        'ID',
        'Id',
        'LinkTitle',
        'SiteURL',
        'Reviewer/Id',
        'Reviewer/Title',
        'Reviewer/EMail',
        'BUH/Id',
        'BUH/Title',
        'BUH/EMail',
        'ProjectManager/Id',
        'ProjectManager/Title',
        'ProjectManager/EMail',
        'TeamSize'
      ];
      const expandFields = ['Reviewer', 'BUH', 'ProjectManager'];
      const filterCandidates = this.deriveSiteUrlCandidates(currentSiteUrlRaw);
      const filterQuery = filterCandidates.length > 0
        ? filterCandidates.map(url => `SiteURL eq '${this.escapeODataString(url)}'`).join(' or ')
        : undefined;

      const fetchOptions = {
        context,
        listTitle: ParentListNames.PPOApprovers,
        select: selectFields,
        expand: expandFields,
        pageSize: 2000,
        filter: undefined as string | undefined
      };

      if (filterQuery) {
        fetchOptions.filter = filterQuery;
      }

      const items = await this.service.fetchAllItems<any>(fetchOptions);

      const buildPersonDetails = (raw: any): { names: string; emails: string; normalizedEmails: string[] } => {
        const entries: Array<{ name: string; email: string }> = [];

        const pushEntry = (person: any): void => {
          if (!person) {
            return;
          }

          if (typeof person === 'string') {
            entries.push({ name: person, email: '' });
            return;
          }

          const name = person?.Title ?? person?.Name ?? '';
          const email = person?.EMail ?? person?.Email ?? person?.email ?? '';

          if (name || email) {
            entries.push({ name, email });
          }
        };

        if (Array.isArray(raw?.results)) {
          raw.results.forEach(pushEntry);
        } else if (Array.isArray(raw)) {
          raw.forEach(pushEntry);
        } else {
          pushEntry(raw);
        }

        const names = entries.map(entry => entry.name).filter(Boolean);
        const emails = entries.map(entry => entry.email).filter(Boolean);

        return {
          names: names.join('; '),
          emails: emails.join('; '),
          normalizedEmails: emails.map(email => this.normalizeEmail(email)).filter(Boolean)
        };
      };

      const normalizedItems: IPPOApprovers[] = [];

      (items || []).forEach((item: any) => {
        const siteUrl: string = item?.SiteURL ?? '';
        const siteMatches = normalizedSiteUrl ? this.normalizeSiteUrl(siteUrl) === normalizedSiteUrl : true;

        if (!siteMatches) {
          return;
        }

        const reviewer = buildPersonDetails(item?.Reviewer);
        const buh = buildPersonDetails(item?.BUH);
        const projectManager = buildPersonDetails(item?.ProjectManager);

        const isCurrentReviewer = currentUserEmailNormalized !== '' && reviewer.normalizedEmails.indexOf(currentUserEmailNormalized) > -1;
        const isCurrentBUH = currentUserEmailNormalized !== '' && buh.normalizedEmails.indexOf(currentUserEmailNormalized) > -1;
        const isCurrentProjectManager = currentUserEmailNormalized !== '' && projectManager.normalizedEmails.indexOf(currentUserEmailNormalized) > -1;

        const currentUserRoles: string[] = [];
        if (isCurrentReviewer) {
          currentUserRoles.push('Reviewer');
        }
        if (isCurrentBUH) {
          currentUserRoles.push('BUH');
        }
        if (isCurrentProjectManager) {
          currentUserRoles.push('ProjectManager');
        }

        normalizedItems.push({
          ID: typeof item?.ID === 'number' ? item.ID : (typeof item?.Id === 'number' ? item.Id : 0),
          LinkTitle: item?.LinkTitle || '',
          teamSize:item?.TeamSize,
          SiteURL: siteUrl,
          Reviewer: reviewer.names,
          ReviewerEmail: reviewer.emails,
          BUH: buh.names,
          BUHEmail: buh.emails,
          ProjectManager: projectManager.names,
          ProjectManagerEmail: projectManager.emails,
          isCurrentReviewer,
          isCurrentBUH,
          isCurrentProjectManager,
          currentUserRoles,
        });
      });

      this.cache = normalizedItems;
      this.cacheTimestamp = now;
      this.cacheKey = cacheKey;

      return this.cache;
    } catch (error: any) {
      throw new Error(ErrorMessages.FAILED_TO_FETCH_PPOAPPROVERS + (error?.message || error));
    }
  }

  public refresh(): void {
    this.cache = null;
    this.cacheTimestamp = 0;
    this.cacheKey = null;
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

const defaultInstance = new PPOApproversRepository();

export default defaultInstance;
export const PPOApproversRepo = defaultInstance;
export const getAllItems = async (useCache: boolean = true, context?: WebPartContext): Promise<IPPOApprovers[]> => defaultInstance.getAllItems(useCache, context);
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();