import { ENV_CONFIG } from '@/utils/environmentConfig';
import { AuthService } from '@/utils/authUtils';

export interface O365User {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  businessPhones: string[];
  mobilePhone?: string;
  assignedLicenses: any[];
  accountEnabled: boolean;
  createdDateTime: string;
  lastSignInDateTime?: string;
}

export interface O365Group {
  id: string;
  displayName: string;
  description?: string;
  groupTypes: string[];
  mail?: string;
  memberCount?: number;
  visibility: string;
  createdDateTime: string;
}

export interface O365Site {
  id: string;
  displayName: string;
  name: string;
  webUrl: string;
  description?: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  siteCollection?: {
    hostname: string;
  };
}

export interface O365License {
  skuId: string;
  skuPartNumber: string;
  servicePlans: Array<{
    servicePlanId: string;
    servicePlanName: string;
    provisioningStatus: string;
  }>;
  assignedLicenses: number;
  availableLicenses: number;
}

export interface O365TenantInfo {
  id: string;
  displayName: string;
  verifiedDomains: Array<{
    name: string;
    isDefault: boolean;
    capabilities: string;
  }>;
  assignedPlans: any[];
  createdDateTime: string;
}

export interface O365Mail {
  id: string;
  subject: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  receivedDateTime: string;
  isRead: boolean;
  importance: string;
  hasAttachments: boolean;
}

export interface O365Team {
  id: string;
  displayName: string;
  description?: string;
  visibility: string;
  isArchived: boolean;
  createdDateTime: string;
  memberCount?: number;
  channels?: O365Channel[];
}

export interface O365Channel {
  id: string;
  displayName: string;
  description?: string;
  membershipType: string;
  createdDateTime: string;
}

export interface O365Permission {
  id: string;
  principalName: string;
  principalType: string;
  roleDefinition: string;
  directoryScope: string;
  isActive: boolean;
}

export class GraphApiService {
  private static baseUrl = ENV_CONFIG.GRAPH_API_BASE_URL;
  private static betaUrl = ENV_CONFIG.GRAPH_API_BETA_URL;

  private static async makeRequest<T>(endpoint: string, useBeta: boolean = false): Promise<T> {
    const token = await AuthService.getGraphToken();
    if (!token) {
      throw new Error('No valid authentication token available');
    }

    const baseUrl = useBeta ? this.betaUrl : this.baseUrl;
    const response = await fetch(`${baseUrl}${endpoint}`, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({ error: 'Unknown error' }));
      throw new Error(`Graph API error: ${response.status} - ${error.error?.message || response.statusText}`);
    }

    return response.json();
  }

  static async getUsers(top: number = 100): Promise<{ users: O365User[]; nextLink?: string }> {
    try {
      const response = await this.makeRequest<{ value: O365User[]; '@odata.nextLink'?: string }>(`/users?$top=${top}&$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,businessPhones,mobilePhone,assignedLicenses,accountEnabled,createdDateTime,signInActivity`);
      
      return {
        users: response.value,
        nextLink: response['@odata.nextLink']
      };
    } catch (error) {
      console.error('Error fetching users:', error);
      throw error;
    }
  }

  static async getUserById(userId: string): Promise<O365User> {
    return this.makeRequest<O365User>(`/users/${userId}`);
  }

  static async getCurrentUser(): Promise<O365User> {
    return this.makeRequest<O365User>('/me');
  }

  static async getGroups(): Promise<O365Group[]> {
    try {
      const response = await this.makeRequest<{ value: O365Group[] }>('/groups?$select=id,displayName,description,groupTypes,mail,visibility,createdDateTime');
      return response.value;
    } catch (error) {
      console.error('Error fetching groups:', error);
      throw error;
    }
  }

  static async getGroupMembers(groupId: string): Promise<O365User[]> {
    try {
      const response = await this.makeRequest<{ value: O365User[] }>(`/groups/${groupId}/members?$select=id,displayName,mail,userPrincipalName`);
      return response.value;
    } catch (error) {
      console.error('Error fetching group members:', error);
      throw error;
    }
  }

  static async getSites(): Promise<O365Site[]> {
    try {
      const response = await this.makeRequest<{ value: O365Site[] }>('/sites?search=*&$select=id,displayName,name,webUrl,description,createdDateTime,lastModifiedDateTime,siteCollection');
      return response.value;
    } catch (error) {
      console.error('Error fetching sites:', error);
      throw error;
    }
  }

  static async getSiteById(siteId: string): Promise<O365Site> {
    return this.makeRequest<O365Site>(`/sites/${siteId}`);
  }

  static async getSubscribedSkus(): Promise<O365License[]> {
    try {
      const response = await this.makeRequest<{ value: any[] }>('/subscribedSkus');
      
      return response.value.map(sku => ({
        skuId: sku.skuId,
        skuPartNumber: sku.skuPartNumber,
        servicePlans: sku.servicePlans || [],
        assignedLicenses: sku.consumedUnits || 0,
        availableLicenses: (sku.prepaidUnits?.enabled || 0) - (sku.consumedUnits || 0)
      }));
    } catch (error) {
      console.error('Error fetching licenses:', error);
      throw error;
    }
  }

  static async getTenantInfo(): Promise<O365TenantInfo> {
    try {
      const response = await this.makeRequest<O365TenantInfo>('/organization?$select=id,displayName,verifiedDomains,assignedPlans,createdDateTime');
      return Array.isArray(response) ? response[0] : response;
    } catch (error) {
      console.error('Error fetching tenant info:', error);
      throw error;
    }
  }

  static async getMailMessages(top: number = 50): Promise<O365Mail[]> {
    try {
      const response = await this.makeRequest<{ value: O365Mail[] }>(`/me/messages?$top=${top}&$select=id,subject,from,receivedDateTime,isRead,importance,hasAttachments&$orderby=receivedDateTime desc`);
      return response.value;
    } catch (error) {
      console.error('Error fetching mail messages:', error);
      throw error;
    }
  }

  static async getUnreadMailCount(): Promise<number> {
    try {
      const response = await this.makeRequest<{ value: number }>('/me/mailFolders/inbox/messages/$count?$filter=isRead eq false');
      return response.value;
    } catch (error) {
      console.error('Error fetching unread mail count:', error);
      return 0;
    }
  }

  static async getMailFolders(): Promise<any[]> {
    try {
      const response = await this.makeRequest<{ value: any[] }>('/me/mailFolders?$select=id,displayName,childFolderCount,unreadItemCount,totalItemCount');
      return response.value;
    } catch (error) {
      console.error('Error fetching mail folders:', error);
      throw error;
    }
  }

  static async getJoinedTeams(): Promise<O365Team[]> {
    try {
      const response = await this.makeRequest<{ value: O365Team[] }>('/me/joinedTeams?$select=id,displayName,description,visibility,isArchived,createdDateTime');
      return response.value;
    } catch (error) {
      console.error('Error fetching joined teams:', error);
      throw error;
    }
  }

  static async getTeamChannels(teamId: string): Promise<O365Channel[]> {
    try {
      const response = await this.makeRequest<{ value: O365Channel[] }>(`/teams/${teamId}/channels`);
      return response.value;
    } catch (error) {
      console.error('Error fetching team channels:', error);
      throw error;
    }
  }

  static async getTeamMembers(teamId: string): Promise<any[]> {
    try {
      const response = await this.makeRequest<{ value: any[] }>(`/teams/${teamId}/members`);
      return response.value;
    } catch (error) {
      console.error('Error fetching team members:', error);
      throw error;
    }
  }

  static async getUserRoleAssignments(userId?: string): Promise<O365Permission[]> {
    try {
      const endpoint = userId 
        ? `/directoryObjects/${userId}/memberOf?$select=id,displayName,description`
        : '/me/memberOf?$select=id,displayName,description';
      
      const response = await this.makeRequest<{ value: any[] }>(endpoint);
      
      return response.value.map(role => ({
        id: role.id,
        principalName: role.displayName,
        principalType: 'Group',
        roleDefinition: role.description || 'Member',
        directoryScope: '/',
        isActive: true
      }));
    } catch (error) {
      console.error('Error fetching user role assignments:', error);
      throw error;
    }
  }

  static async getDirectoryRoles(): Promise<any[]> {
    try {
      const response = await this.makeRequest<{ value: any[] }>('/directoryRoles?$select=id,displayName,description,roleTemplateId');
      return response.value;
    } catch (error) {
      console.error('Error fetching directory roles:', error);
      throw error;
    }
  }

  static async getRoleAssignments(): Promise<O365Permission[]> {
    try {
      const response = await this.makeRequest<{ value: any[] }>('/roleManagement/directory/roleAssignments?$expand=principal,roleDefinition', true);
      
      return response.value.map(assignment => ({
        id: assignment.id,
        principalName: assignment.principal?.displayName || 'Unknown',
        principalType: assignment.principal?.['@odata.type']?.split('.').pop() || 'Unknown',
        roleDefinition: assignment.roleDefinition?.displayName || 'Unknown Role',
        directoryScope: assignment.directoryScopeId || '/',
        isActive: true
      }));
    } catch (error) {
      console.error('Error fetching role assignments:', error);
      throw error;
    }
  }

  static async syncAllO365Data(): Promise<{
    users: O365User[];
    groups: O365Group[];
    sites: O365Site[];
    licenses: O365License[];
    tenantInfo: O365TenantInfo;
    mail: O365Mail[];
    teams: O365Team[];
    permissions: O365Permission[];
    mailFolders: any[];
    unreadMailCount: number;
  }> {
    try {
      console.log('Starting comprehensive O365 data sync...');
      
      const [
        usersData,
        groups,
        sites, 
        licenses,
        tenantInfo,
        mail,
        teams,
        permissions,
        mailFolders,
        unreadMailCount
      ] = await Promise.all([
        this.getUsers(500),
        this.getGroups(),
        this.getSites(),
        this.getSubscribedSkus(),
        this.getTenantInfo(),
        this.getMailMessages(100),
        this.getJoinedTeams(),
        this.getUserRoleAssignments(),
        this.getMailFolders(),
        this.getUnreadMailCount()
      ]);

      const teamsWithChannels = await Promise.all(
        teams.map(async (team) => {
          try {
            const channels = await this.getTeamChannels(team.id);
            const members = await this.getTeamMembers(team.id);
            return {
              ...team,
              channels,
              memberCount: members.length
            };
          } catch (error) {
            console.warn(`Failed to get channels for team ${team.id}:`, error);
            return team;
          }
        })
      );

      console.log('O365 comprehensive sync completed:', {
        users: usersData.users.length,
        groups: groups.length,
        sites: sites.length,
        licenses: licenses.length,
        mail: mail.length,
        teams: teamsWithChannels.length,
        permissions: permissions.length,
        mailFolders: mailFolders.length,
        unreadMail: unreadMailCount
      });

      return {
        users: usersData.users,
        groups,
        sites,
        licenses,
        tenantInfo,
        mail,
        teams: teamsWithChannels,
        permissions,
        mailFolders,
        unreadMailCount
      };
    } catch (error) {
      console.error('Error in comprehensive O365 sync:', error);
      throw error;
    }
  }

  static async healthCheck(): Promise<{ status: 'healthy' | 'unhealthy'; details: string }> {
    try {
      await this.makeRequest('/me');
      return { status: 'healthy', details: 'Graph API connection successful' };
    } catch (error) {
      return { 
        status: 'unhealthy', 
        details: `Graph API connection failed: ${error instanceof Error ? error.message : 'Unknown error'}` 
      };
    }
  }
}
