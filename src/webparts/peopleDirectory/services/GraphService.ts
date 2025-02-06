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

        // Prepare batch requests for user presence and photos
        const presenceRequests = usersResponse.value.map(async (user: IUser) => {
            try {
                const presenceResponse = await client.api(`/users/${user.id}/presence`).version('beta').get();
                
                let availability = presenceResponse.availability;
                let activity = presenceResponse.activity;
                let statusMessage = presenceResponse.statusMessage?.message?.content || "";
                let outOfOffice = presenceResponse.outOfOfficeSettings?.isOutOfOffice;

                if (outOfOffice) {
                    availability = "Out of Office";
                    activity = "Out of Office";
                    statusMessage = presenceResponse.outOfOfficeSettings?.message || "";
                }

                let photoUrl = "";
                try {
                    const photoResponse = await client.api(`/users/${user.id}/photos/48x48/$value`).get();
                    photoUrl = URL.createObjectURL(await photoResponse);
                    console.log("photoResponse is " + photoResponse + " photoUrl is " + photoUrl)
                } catch (photoError) {
                    Log.warn("GraphService - Photo", `Could not fetch photo for ${user.displayName}`);
                }
                if (presenceResponse.availability !== "PresenceUnknown") {
                  return {
                      id: user.id,
                      displayName: user.displayName,
                      jobTitle: user.jobTitle,
                      department: user.department,
                      availability,
                      activity,
                      statusMessage,
                      photo: photoUrl || "/_layouts/15/userphoto.aspx?size=S&accountname=" + user.id // Fallback to SharePoint user photo
                  };
                } 
                return null // push nothing if "PresenceUnknown" as this usually means a leaver
            } catch (error) {
                Log.error("GraphService - Presence", error);
                return null;
            }
        });

        const presenceResults = await Promise.all(presenceRequests);
        users.push(...presenceResults.filter((user): user is IUser => user !== null));

    } catch (error) {
        Log.error("GraphService - General", error);
    }

    return users;
  }
}
