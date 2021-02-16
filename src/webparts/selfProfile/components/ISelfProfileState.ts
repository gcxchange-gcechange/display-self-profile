import { IUserInfo } from './IUserInfo';  
  
export interface ISelfProfileState {  
    users: Array<IUserInfo>;  
    userID: string;
    displayName: string;
    mail: string;
    userPrincipalName: string;
    givenName: string;
    surname: string;
    jobTitle: string;
    mobilePhone: string;
    businessPhone: string;
    officeLocation: string;
    streetAddress: string;
    city: string;
    state: string;
    postalCode: string;
    country: string;
    photo: string;
    department: string;
    managerDisplayName: string;
    modalToggle: boolean;
}  