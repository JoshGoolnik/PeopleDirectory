import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUser } from '../models/IUser';
import { Log } from '@microsoft/sp-core-library';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export class GraphService {
  constructor(private context: WebPartContext) {}

  public async getUsersWithPresence(): Promise<IUser[]> {
    const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
    const users: IUser[] = [];

    if (!client) {
      Log.error("GraphService - General", new Error("MSGraphClientV3 is not initialized."));
      return users;
    }

    try {
      const usersResponse = await client
        .api('/users')
        .version('v1.0')
        .header('ConsistencyLevel', 'eventual')
        .filter('Department ne null AND surname ne null')
        .select('id,displayName,jobTitle,department')
        .orderby('displayName')
        .top(50)
        .count(true)
        .get();

      Log.info("GraphService", `Fetched ${usersResponse.value.length} users`);

      // Prepare batch requests for user presence
      const presenceRequests = usersResponse.value.map((user:IUser) => 
        client.api(`/users/${user.id}/presence`).version('beta').get()
          .then(presenceResponse => {
            let availability = presenceResponse.availability;
            let activity = presenceResponse.activity;
            let statusMessage = presenceResponse.statusMessage?.message?.content || "";
            let outOfOffice = presenceResponse.outOfOfficeSettings?.isOutOfOffice;

            Log.info("GraphService", `Fetched ${user.displayName}, out of office status ${outOfOffice}`);
            if (outOfOffice) {
              availability = "Out of Office";
              activity = "Out of Office";
              statusMessage = presenceResponse.outOfOfficeSettings?.message || "";
            }
            // Only push to users if availability is not unknown
            if (presenceResponse.availability !== "PresenceUnknown") {
              return {
                displayName: user.displayName,
                jobTitle: user.jobTitle,
                department: user.department,
                availability,
                activity,
                statusMessage
              };
            }
            return null; // Avoid pushing if presence is unknown
          })
          .catch(presenceError => {
            Log.error("GraphService - Presence", presenceError);
            return null; // Handle errors gracefully
          })
      );

      // Resolve all presence requests concurrently
      const presenceResults = await Promise.all(presenceRequests);
      // Filter out any null results due to errors or unknown presence
      users.push(...presenceResults.filter((user): user is IUser => user !== null));
      
    } catch (error) {
      Log.error("GraphService - General", error);
    }

    return users;
  }
}
