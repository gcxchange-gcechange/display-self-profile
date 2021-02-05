import * as React from 'react';
import { MSGraphClient } from '@microsoft/sp-http';  
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'; 
import {
  autobind,
  PrimaryButton,
  Persona,
  PersonaSize,
  Stack,
  Modal,
  TextField
} from 'office-ui-fabric-react'; 
import styles from './SelfProfile.module.scss';
import { ISelfProfileProps } from './ISelfProfileProps';
import { ISelfProfileState } from './ISelfProfileState';
import { IUserInfo } from './IUserInfo';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SelfProfile extends React.Component<ISelfProfileProps, ISelfProfileState> {
  constructor(props: ISelfProfileProps, state: ISelfProfileState) {  
    super(props);  
  
    // Initialize the state of the component  
    this.state = {  
      users: [], 
      displayName: '',
      mail: '',
      userPrincipalName: '',
      givenName: '',
      surname: '',
      jobTitle: '',
      mobilePhone: '', 
      officeLocation: '',
      photo: '',
      managerDisplayName: '',
      modalToggle: false
    };  
  }

@autobind
private getUserDetails(): void {  
  this.props.context.msGraphClientFactory  
    .getClient()  
    .then((client: MSGraphClient): void => {  
      // Get user information from the Microsoft Graph  
      client  
        .api('/me')
        .version("v1.0")
        .select(["displayName","mail","userPrincipalName","givenName","surname","jobTitle","mobilePhone","officeLocation","photo","department","manager"]) 
        .get((error, result: MicrosoftGraph.User, rawResponse?: any) => {  
          // handle the response  
          if (error) {  
            console.log(error);
            return;  
          }  
          console.log(result);
          // Prepare the output array  
          var users: Array<IUserInfo> = new Array<IUserInfo>();  
          // Map the JSON response to the output array  
          // TODO remove this
          users.push({
            displayName: result.displayName,
            mail: result.mail,
            userPrincipalName: result.userPrincipalName,
            givenName: result.givenName,
            surname: result.surname,
            jobTitle: result.jobTitle,
            mobilePhone: result.mobilePhone,
            officeLocation: result.officeLocation
          });
          // Update the component state accordingly to the result  
          this.setState(  
            {  
              users: users,  
              displayName: result.displayName,
              mail: result.mail,
              givenName: result.givenName,
              surname: result.surname,
              jobTitle: result.jobTitle,
              mobilePhone: result.mobilePhone,
              officeLocation: result.officeLocation
            }  
          );  
        });  
    });  
}

@autobind
private getUserPhoto(): void {  
  this.props.context.msGraphClientFactory  
    .getClient()  
    .then((client: MSGraphClient): void => {  
      // Get user information from the Microsoft Graph  
      client  
        .api('/me/photo')
        .version("v1.0")
        .get((error, result: MicrosoftGraph.User, rawResponse?: any) => {  
          // handle the response  
          if (error) {  
            console.log(error);
            return;  
          }  
          // Log the photo response for now
          // TODO add photo / photo URL to component state
          console.log(result);
        });  
    });  
}

@autobind
private getUserManager(): void {  
  this.props.context.msGraphClientFactory  
    .getClient()  
    .then((client: MSGraphClient): void => {  
      // Get user information from the Microsoft Graph  
      client  
        .api('/me/manager')
        .version("v1.0")
        .get((error, result: MicrosoftGraph.User, rawResponse?: any) => {  
          // handle the response  
          if (error) {  
            console.log(error);
            return;  
          }  
          console.log(result);
          this.setState({
            managerDisplayName: result.displayName
          });
        });  
    });  
}

@autobind
private toggle(): void {
  this.setState({
    modalToggle: !this.state.modalToggle,
  });
}

componentDidMount() {
  // TODO try to grab all data in one function
  this.getUserDetails();
  this.getUserPhoto();
  this.getUserManager();
}

  public render(): React.ReactElement<ISelfProfileProps> {
    return (
      <div className={ styles.selfProfile }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div>
              <Stack horizontalAlign="end">
                <PrimaryButton
                  text="EDIT PH"
                  onClick={this.toggle}
                />
              </Stack>
              <Modal
                isOpen={this.state.modalToggle}
              >
                {
                  // TODO perhaps make modal form its own component
                  // TODO make the form work
                  // TODO i18n
                }
                <div className={ styles.title }>Hello! A form will go here</div>
                <PrimaryButton
                  onClick={this.toggle}
                >
                  CLOSE PH
                </PrimaryButton>
                <div className={ styles.modalContainer }>
                  <TextField
                    label="Name"
                    value={this.state.displayName}
                  />
                  <TextField
                    label="Email"
                    value={this.state.mail}
                  />
                  <Stack horizontal>
                    <TextField
                      label="Job Title EN"
                      className={ styles.formMr }
                      value={this.state.jobTitle}
                    />
                    <TextField
                      label="Job Title FR"
                    />
                  </Stack>
                  <Stack horizontal>
                    <TextField
                      label="Street Address"
                      className={ styles.formMr }
                    />
                    <TextField
                      label="City"
                      className={ styles.formMr }
                    />
                  </Stack>
                  <Stack horizontal>
                    <TextField
                      label="Province"
                      className={ styles.formMr }
                    />
                    <TextField
                      label="Postal Code"
                      className={ styles.formMr }
                    />
                    <TextField
                      label="Country"
                    />
                  </Stack>
                  <Stack horizontal>
                    <TextField
                      label="Mobile Phone"
                      className={ styles.formMr }
                      value={this.state.mobilePhone}
                    />
                    <TextField
                      label="Office Phone"
                    />
                  </Stack>
                  <TextField
                    label="Manager (PH)"
                  />
                </div>
              </Modal>
            </div>
            <div className={ styles.column }>
              {
                // TODO grab photo URL
              }
              <Persona 
                text= {(this.state.displayName) && this.state.displayName}
                secondaryText= {(this.state.jobTitle) ? this.state.jobTitle : 'Job Title PH'}
                tertiaryText= "Department Here"
                size={PersonaSize.size72}
              />
              <div>
                <div>
                  <div className={ styles.dataContainer }>
                    <div className={ styles.dataLabel }>Email</div>
                    {
                      (this.state.mail) ?
                      <div>{this.state.mail}</div> :
                      <div>Mail PH</div>
                    }
                  </div>
                </div>
                <Stack horizontal>
                  <div className={ styles.dataContainer }>
                    <div className={ styles.dataLabel }>Mobile Phone</div>
                    {
                      (this.state.mobilePhone) ?
                      <div>{this.state.mobilePhone}</div> :
                      <div>Phone N/A</div>
                    }
                  </div>
                  <div className={ styles.dataContainer }>
                    <div className={ styles.dataLabel }>Office Phone</div>
                    <div>TEST 2</div>
                  </div>
                  <div className={ styles.dataContainer }>
                    <div className={ styles.dataLabel }>Office Location</div>
                    {
                      // TODO Change this to address
                      (this.state.officeLocation) ?
                      <div>{this.state.officeLocation}</div> :
                      <div>N/A</div>
                    }
                  </div>
                </Stack>
                <div className={ styles.dataContainer }>
                  <div className={ styles.dataLabel }>Manager</div>  
                  {
                    (this.state.managerDisplayName) ?
                    <div>{this.state.managerDisplayName}</div> :
                    <div>Manager N/A</div>
                  }
                </div>
              </div>
            </div>
          </div>
          {
            // this is a button for testing API call
          }
            <PrimaryButton
              text='TEST BTN'
              title='TEST'
              onClick={this.getUserDetails}
            />
        </div>
      </div>
    );
  }
}
