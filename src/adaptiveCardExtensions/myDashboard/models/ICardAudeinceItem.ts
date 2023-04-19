export interface ICardAudienceItem {
    title: string,
    identifier: string,
    groupName: string,
    groupId: string,
    locked: boolean,
    // Is the user subscribed to this audience
    isMember: string
}