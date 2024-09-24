import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUser } from '../models/IUser';
import { Log } from '@microsoft/sp-core-library';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export class GraphService {
  constructor(private context: WebPartContext) {    
  }

  public async getUsersWithPresence(): Promise<IUser[]> {
    const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
    const users: IUser[] = [];

    if (!client) {
      Log.error("GraphService - General", new Error("MSGraphClientV3 is not initialized."));
      return users;
    }

    try {
      // Use MSGraphClientV3 to fetch users
      const usersResponse = await client
        .api('/users')
        .version('v1.0')
        .header('ConsistencyLevel','eventual')
        .filter('officeLocation ne null AND Department ne null AND surname ne null')
        .select('id,displayName,jobTitle,department')
        .orderby('displayName')
        .top(50)
        .count(true)
        .get();
        
      Log.info("GraphService", `Fetched ${usersResponse.value.length} users`);

      // Fetch presence for each user
      for (let user of usersResponse.value) {
        try {
          const presenceResponse = await client
            .api(`/users/${user.id}/presence`)
            .version('beta')
            .get();

          if (presenceResponse.availability != "PrescenceUnknown")
          {
            users.push({
              displayName: user.displayName,
              jobTitle: user.jobTitle,
              department: user.department,
              availability: presenceResponse.availability,
              activity: presenceResponse.activity,
              statusMessage: presenceResponse.statusMessage?.message?.content || "",
              workLocation: user.officeLocation
            });
          }
        } catch (presenceError) {
          Log.error("GraphService - Presence", presenceError);  // Log presence API errors
        }
      }
    } catch (error) {
      Log.error("GraphService - General", error);  // Log general Graph API errors
    }

    return users;
  }
}
