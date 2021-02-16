import * as React from 'react';
import { MSGraphClient, IHttpClientOptions, AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';  
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'; 
import {
  autobind,
  PrimaryButton,
  Button,
  Persona,
  PersonaSize,
  Stack,
  Modal,
  TextField,
  MaskedTextField
} from 'office-ui-fabric-react'; 
import styles from './SelfProfile.module.scss';
import { ISelfProfileProps } from './ISelfProfileProps';
import { ISelfProfileState } from './ISelfProfileState';
import { IUserInfo } from './IUserInfo';
import EditModal from './EditModal';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SelfProfile extends React.Component<ISelfProfileProps, ISelfProfileState> {
  constructor(props: ISelfProfileProps, state: ISelfProfileState) {  
    super(props);  
  
    // Initialize the state of the component  
    this.state = {  
      users: [], 
      userID: '',
      displayName: '',
      mail: '',
      userPrincipalName: '',
      givenName: '',
      surname: '',
      jobTitle: '',
      mobilePhone: '',
      businessPhone: '',
      officeLocation: '',
      streetAddress: '',
      city: '',
      state: '',
      postalCode: '',
      country: '',
      photo: '',
      department: '',
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
      // https://graph.microsoft.com/v1.0/me/?$select=displayName,country,streetAddress,postalCode,state,city,photo,mail,userPrincipalName,givenName,surname,jobTitle,mobilePhone,businessPhones,department,officeLocation&$expand=manager

      client  
        .api('/me')
        .version("v1.0")
        .select(["id,displayName","mail","userPrincipalName","givenName","surname","jobTitle","mobilePhone","officeLocation","department","streetAddress","city","state","postalCode","country","businessPhones"]) 
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
              userID: result.id,
              displayName: result.displayName,
              mail: result.mail,
              givenName: result.givenName,
              surname: result.surname,
              jobTitle: result.jobTitle,
              mobilePhone: result.mobilePhone,
              businessPhone: result.businessPhones[0],
              officeLocation: result.officeLocation,
              streetAddress: result.streetAddress,
              city: result.city,
              state: result.state,
              postalCode: result.postalCode,
              country: result.country,
              department: result.department,
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
      // https://graph.microsoft.com/v1.0/me/manager?$select=displayName
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

@autobind
private closeModal(): void {
  this.getUserDetails();
  this.setState({
    modalToggle: false
  });
}

@autobind
private sendUserData(): void {
  // TODO add the call the function all here
  console.log("You clicked the save button!");
  console.log(this.state);
  const reqHeaders: Headers = new Headers();
  reqHeaders.append('Content-type', 'application/json');
  const reqBody: string = JSON.stringify({
    //TODO change userID Back to state
    'user': {
      'userID': this.state.userID,
      'jobTitle': this.state.jobTitle,
      'firstName': this.state.givenName,
      'lastName': this.state.surname,
      'mobilePhone': this.state.mobilePhone,
      'streetAddress': this.state.streetAddress,
      'city': this.state.city,
      'province': this.state.state,
      'postalcode': this.state.postalCode,
      'country': this.state.country,
      'department': this.state.department,
      'businessPhones': this.state.businessPhone,
    }
  });
  const options: IHttpClientOptions = {
    headers: reqHeaders,
    body: reqBody
  }
  console.log(options);
  this.props.context.aadHttpClientFactory
      // Add Client
      .getClient('')
      .then((client: AadHttpClient): void => {
        client
          // Add URL
          .post('', AadHttpClient.configurations.v1, options)
          .then((response: HttpClientResponse) => {
            console.log(response);
            return response.json();
          })
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
                <div className={ styles.modalContainer }>
                  <div className={ styles.title }>Edit</div>
                  <TextField
                    label="First Name"
                    value={this.state.givenName}
                    onChange={(e) => {
                      this.setState({
                        givenName: (e.target as HTMLInputElement).value
                      });
                    }}
                  />
                  <TextField
                    label="Last Name"
                    value={this.state.surname}
                    onChange={(e) => {
                      this.setState({
                        surname: (e.target as HTMLInputElement).value
                      });
                    }}
                  />
                  <TextField
                    label="Email"
                    value={this.state.mail}
                    onChange={(e) => {
                      this.setState({
                        mail: (e.target as HTMLInputElement).value
                      });
                    }}
                  />
                  <Stack horizontal>
                    <TextField
                      label="Job Title EN"
                      className={ styles.formMr }
                      value={this.state.jobTitle}
                      onChange={(e) => {
                        this.setState({
                          jobTitle: (e.target as HTMLInputElement).value
                        });
                      }}
                    />
                    <TextField
                      label="Job Title FR"
                    />
                  </Stack>
                  <Stack horizontal>
                    <TextField
                      label="Street Address"
                      className={ styles.formMr }
                      value={this.state.streetAddress}
                      onChange={(e) => {
                        this.setState({
                          streetAddress: (e.target as HTMLInputElement).value
                        });
                      }}
                    />
                    <TextField
                      label="City"
                      className={ styles.formMr }
                      value={this.state.city}
                      onChange={(e) => {
                        this.setState({
                          city: (e.target as HTMLInputElement).value
                        });
                      }}
                    />
                  </Stack>
                  <Stack horizontal>
                    <TextField
                      label="Province"
                      className={ styles.formMr }
                      value={this.state.state}
                      onChange={(e) => {
                        this.setState({
                          state: (e.target as HTMLInputElement).value
                        });
                      }}
                    />
                    <TextField
                      label="Postal Code"
                      className={ styles.formMr }
                      value={this.state.postalCode}
                      onChange={(e) => {
                        this.setState({
                          postalCode: (e.target as HTMLInputElement).value
                        });
                      }}
                    />
                    <TextField
                      label="Country"
                      value={this.state.country}
                      onChange={(e) => {
                        this.setState({
                          country: (e.target as HTMLInputElement).value
                        });
                      }}
                    />
                  </Stack>
                  <Stack horizontal>
                    <MaskedTextField
                      label="Mobile Phone"
                      mask="(999) 999 - 9999"
                      className={ styles.formMr }
                      value={(this.state.mobilePhone) ? this.state.mobilePhone : ''}
                      onChange={(e) => {
                        this.setState({
                          mobilePhone: (e.target as HTMLInputElement).value
                        });
                      }}
                    />
                    <MaskedTextField
                      label="Office Phone"
                      mask="(999) 999 - 9999"
                      value={this.state.businessPhone}
                      onChange={(e) => {
                        this.setState({
                          businessPhone: (e.target as HTMLInputElement).value
                        });
                      }}
                    />
                  </Stack>
                  <TextField
                    label="Manager (PH)"
                    disabled
                  />
                  <TextField
                    label="Department"
                    disabled
                  />
                  <div className={ styles.actionBtnContainer }>
                    <Stack horizontal horizontalAlign="end">
                      <Button
                        onClick={this.closeModal}
                      >
                        CLOSE PH
                      </Button>
                      <PrimaryButton
                        text="SAVE PH"
                        onClick={this.sendUserData}
                      />  
                    </Stack>
                  </div>
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
                tertiaryText= {(this.state.department) ? this.state.department : 'Department PH'}
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
                    {
                      (this.state.businessPhone) ?
                      <div>{this.state.businessPhone}</div> :
                      <div>N/A</div>
                    }
                  </div>
                  <div className={ styles.dataContainer }>
                    <div className={ styles.dataLabel }>Office Location</div>
                    {
                      (this.state.officeLocation) &&
                      <div>{this.state.officeLocation}</div>
                    }
                    {
                      (this.state.streetAddress) &&
                      <span>{this.state.streetAddress} </span>
                    }
                    {
                      (this.state.city) &&
                      <span>{this.state.city} </span>
                    }
                    {
                      (this.state.state) &&
                      <span>{this.state.state} </span>
                    }
                    {
                      (this.state.postalCode) &&
                      <span>{this.state.postalCode} </span>
                    }
                    {
                      (this.state.country) &&
                      <span>{this.state.country} </span>
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
        </div>
      </div>
    );
  }
}
