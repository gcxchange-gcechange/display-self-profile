import * as React from 'react';
import {
  autobind,
  PrimaryButton,
  Stack,
  Modal,
  TextField
} from 'office-ui-fabric-react'; 
import styles from './SelfProfile.module.scss';
import { IEditModalProps } from './IEditModalProps';
import { ISelfProfileState } from './ISelfProfileState';

export default class EditModal extends React.Component<IEditModalProps, ISelfProfileState> {
  constructor(props: IEditModalProps, state: ISelfProfileState) {  
    super(props);  
  
    // Initialize the state of the component  
    this.state = {  
      users: [], 
      userID: '',
      displayName: this.props.displayName,
      mail: this.props.mail,
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

  /*
  static getDerivedStateFromProps(nextProps, prevState) {
    console.log('FIRE');
    if (prevState.displayName !== nextProps.displayName) {
      return { displayName: nextProps.displayName };
    }

    return null;
}*/

  @autobind
  private sendUserData(): void {
    // TODO add the call the function all here
    console.log("You clicked the save button!");
    console.log(this.state);
    console.log(this.props.displayName);
    console.log(this.props.mail);
  }

  @autobind
  private toggle(): void {
    this.setState({
      modalToggle: !this.state.modalToggle,
    });
  }

  public render(): React.ReactElement<IEditModalProps> {
    return (
      <div>
        I am the modal component
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
                    onChange={(e) => {
                      console.log((e.target as HTMLInputElement).value);
                      this.setState({
                        displayName: (e.target as HTMLInputElement).value
                      });
                    }}
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
                      value={this.state.streetAddress}
                    />
                    <TextField
                      label="City"
                      className={ styles.formMr }
                      value={this.state.city}
                    />
                  </Stack>
                  <Stack horizontal>
                    <TextField
                      label="Province"
                      className={ styles.formMr }
                      value={this.state.state}
                    />
                    <TextField
                      label="Postal Code"
                      className={ styles.formMr }
                      value={this.state.postalCode}
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
                  <div>
                    <PrimaryButton
                      text="SAVE PH"
                      onClick={this.sendUserData}
                    />
                  </div>
                </div>
              </Modal>
            </div>
      </div>
    )}

}