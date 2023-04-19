import { MSGraphClientV3 } from '@microsoft/sp-http'
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphService {
    Init: (graphClient: MSGraphClientV3, loggedinUsersUPN: string) => void;
    GetLoggedInUsersDirectoryId: () => Promise<void>;    
    GetGroupWithName: (groupName: string) => Promise<MicrosoftGraph.Group>;
    IsLoggedInUserAMember: (groupId: string) => Promise<boolean>;
    AddLoggedInUserToGroup: (groupId: string) => Promise<boolean>;
    RemoveLoggedInUserFromGroup: (groupId: string) => Promise<boolean>;
}



export class GraphService implements IGraphService {
    
    
    private graphClient: MSGraphClientV3;
    private userDirectoryId: string
    
    public Init(graphClient: MSGraphClientV3, loggedinUsersDirectoryId: string): void {
        this.graphClient = graphClient
        this.userDirectoryId = loggedinUsersDirectoryId
    }

    // may not be needed
    public async GetLoggedInUsersDirectoryId(): Promise<void> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }

        var user = await this.graphClient
            .api('/me')
            .get();

        this.userDirectoryId = user.id;
        
        return Promise.resolve();
    }

    public async GetGroupWithName(groupName: string): Promise<MicrosoftGraph.Group> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }

        try {

            var groups: any = await this.graphClient
                .api('/groups')
                .filter(`displayName eq \'${groupName}\'`)
                .get();

        return groups.value[0];
        } catch (error : any) {
            console.log(error)
            return undefined;
        }
    }

    

    public async AddLoggedInUserToGroup(groupId: string): Promise<boolean> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }

        try { 
        var response = await this.graphClient
            .api(`/groups/${groupId}/members/$ref`)
            .post({
                "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${this.userDirectoryId}`
            })
            return true;
        } catch (error) {
            return false;
        }
    }

    public async RemoveLoggedInUserFromGroup(groupId: string): Promise<boolean> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }

        try {
        var response = await this.graphClient
            .api(`/groups/${groupId}/members/${this.userDirectoryId}/$ref`)
            .delete()
            return true;
        } catch (error) {
            return false;
        }
    }

    public async IsLoggedInUserAMember(groupId: string): Promise<boolean> {
        if (this.graphClient === undefined){
            throw new Error('GraphService not initialized!')
        }

        try {
            var memberOfResponse = await this.graphClient.api(`/users/${this.userDirectoryId}/memberOf`)
                .filter(`id eq \'${groupId}\'`)
                .get();

            return true

        } catch (error) {
            // likey 404 returned from api
            return false;
        }
    }

}


export const graphService = new GraphService();