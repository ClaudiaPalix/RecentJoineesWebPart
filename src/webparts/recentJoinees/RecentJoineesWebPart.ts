import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneDropdownOptionType,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './RecentJoineesWebPart.module.scss';
import * as strings from 'RecentJoineesWebPartStrings';
import iconPerson from './assets/iconPerson.png';

interface ListItem {
  UserName: string;
  JoiningDate: string;
  Department: string;
  Company: string;
  ProfilePicture: {
    Url: string;
  }
  Email: string;
  RefreshedOn: string;
}

export interface IRecentJoineesWebPartProps {
  description: string;
  message: string;
  UserName: string;
  JoiningDate: string;
  Department: string;
  Company: string;
  ProfilePicture: {
    Url: string;
  }
  Email: string;
  RefreshedOn: string;
}

export default class RecentJoineesWebPart extends BaseClientSideWebPart<IRecentJoineesWebPartProps> {
  private userEmail: string = "";

  private async userDetails(): Promise<void> {
    // Ensure that you have access to the SPHttpClient
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
  
    // Use try-catch to handle errors
    try {
      // Get the current user's information
      const response: SPHttpClientResponse = await spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1);
      const userProperties: any = await response.json();
  
      console.log("User Details:", userProperties);
  
      // Access the userPrincipalName from userProperties
      const userPrincipalNameProperty = userProperties.UserProfileProperties.find((property: any) => property.Key === 'SPS-UserPrincipalName');
  
      if (userPrincipalNameProperty) {
        this.userEmail = userPrincipalNameProperty.Value.toLowerCase();
        // console.log('User Email using User Principal Name:', this.userEmail);
        // Now you can use this.userEmail as needed
      } else {
        console.error('User Principal Name not found in user properties');
      }
    } catch (error) {
      console.error('Error fetching user properties:', error);
    }
  } 

  public getItemsFromSPList(listName: string): Promise<any[]> {
    return new Promise((resolve, reject) => {
      let open = indexedDB.open("MyDatabase", 1);
   
      open.onsuccess = function() {
        let db = open.result;
        let tx = db.transaction(`${listName}`, "readonly");
        let store = tx.objectStore(`${listName}`);
   
        let getAllRequest = store.getAll();
   
        getAllRequest.onsuccess = function() {
          resolve(getAllRequest.result);
        };
   
        getAllRequest.onerror = function() {
          reject(getAllRequest.error);
        };
      };
   
      open.onerror = function() {
        reject(open.error);
      };
    });
  }

  public render(): void {
   this.userDetails().then(() => {
    const decodedDescription = decodeURIComponent(this.properties.description);
    // console.log("Title: ",decodedDescription);
    this.domElement.innerHTML = `
      <div class="${styles.parentDiv}">
        <div id="upcomingBirthdays" class="${styles.upcomingBirthdays}">
          <h3>${decodedDescription}</h3>
        </div>
      </div>`;
      this._renderButtons();
    });
  }

  private _renderButtons(): void {
    const buttonsContainer: HTMLElement | null = this.domElement.querySelector('#upcomingBirthdays');

    // console.log("User's Email from LoginName: ", this.userEmail);
    const adminEmailSplit: string[] = this.userEmail.split('.admin@');
    if (this.userEmail.includes(".admin@")){
        console.log("Admin Email after split: ", adminEmailSplit);
    }
    const parts = this.userEmail.split('_');
    const secondPart = parts.length > 1 ? parts[1] : '';
    const otherUsersSplit =  secondPart.split('.com')[0];
    if (this.userEmail.includes("_")){
        console.log("User's company after split: ", otherUsersSplit);
    }

    this.getItemsFromSPList("SPList")
    .then((items: ListItem[]) => {
      console.log("All items retrieved:", items);
        let buttonsCreated = 0; // Variable to keep track of the number of buttons created
        if (items && items.length > 0) {
          // Set time to 23:59:59 for today
          const today = new Date();
          today.setHours(23, 59, 59, 999);
        
          // Set time to 00:00:00 for thirtyDaysEarlier
          const thirtyDaysEarlier = new Date(today);
          thirtyDaysEarlier.setDate(today.getDate() - 29);
          thirtyDaysEarlier.setHours(0, 0, 0, 0);
          // console.log("today: ", today);
          // console.log("thirtyDaysEarlier: ", thirtyDaysEarlier);
        
          const filteredItemsWithDate = items.filter(item => {
            if (!item.JoiningDate || !item.UserName) {
              return false;
            }
        
            let itemDate = this.adjustDateForTimeZone(item.JoiningDate);
            itemDate.setHours(0, 0, 0, 0); // Set time to 00:00:00
            // console.log("joiningDate for user", item.UserName, "is: ", itemDate);
            // console.log("joiningDate <= today:", itemDate <= today);
            // console.log("joiningDate >= thirtyDaysEarlier:", itemDate >= thirtyDaysEarlier);
            return itemDate <= today && itemDate >= thirtyDaysEarlier;
          });

        filteredItemsWithDate.sort((a, b) => new Date(b.JoiningDate).getTime() - new Date(a.JoiningDate).getTime());

        
        console.log('filteredItems: ',filteredItemsWithDate);

        filteredItemsWithDate.forEach((item: IRecentJoineesWebPartProps) => {

            if(!item.Company){
              item.Company = " ";
            }
            if((this.userEmail.includes("@"+item.Company.toLowerCase()+".") && !this.userEmail.includes(".admin@") && !otherUsersSplit) || (this.userEmail.includes(".admin@") && adminEmailSplit.includes("@"+item.Company.toLowerCase()+".")) || (otherUsersSplit.length >= 0 && otherUsersSplit.includes(item.Company.toLowerCase()))){
                    // console.log("Creating button for ", item.UserName);
                    const buttonDiv: HTMLDivElement = document.createElement('div');
                    buttonDiv.classList.add(styles.innerContents);

                    const profileSection: HTMLDivElement = document.createElement('div');
                    profileSection.classList.add(styles.profileSection); 

                    const imgBox: HTMLDivElement = document.createElement('div');
                    imgBox.classList.add(styles.imgBox); 
                    const img: HTMLImageElement = document.createElement('img');
                    img.src = item.ProfilePicture && item.ProfilePicture.Url ? item.ProfilePicture.Url : iconPerson;

                    // Add an error event listener to handle image load errors
                    img.addEventListener('error', () => {
                      img.src = iconPerson; // Set to default image if an error occurs
                    });

                    imgBox.appendChild(img);
                    
                    const nameDiv: HTMLDivElement = document.createElement('div');
                    nameDiv.classList.add(styles.name); 
                    const h5: HTMLHeadingElement = document.createElement('h5');
                    h5.textContent = item.UserName;
                    const spanCompany: HTMLSpanElement = document.createElement('span');
                    spanCompany.textContent = item.Company;
                    const spanDept: HTMLSpanElement = document.createElement('span');
                    spanDept.textContent = item.Department;

                    const birthdayMonthDay = this.adjustDateForTimeZone(item.JoiningDate).toISOString().substring(5, 10);
                    const formattedBirthday = this.formatBirthday(birthdayMonthDay);
                    // console.log(this.formatBirthday(formattedBirthday))
                    const birthdayText = `Joining Date: ${formattedBirthday}`;

                    const birthdayElement: HTMLSpanElement = document.createElement('span');
                    birthdayElement.textContent = birthdayText;

                    nameDiv.appendChild(h5);  
                    nameDiv.appendChild(spanDept);
                    nameDiv.appendChild(birthdayElement);
                    profileSection.appendChild(imgBox); 
                    profileSection.appendChild(nameDiv);

                    const chatBtn: HTMLButtonElement = document.createElement('button');
                    chatBtn.classList.add(styles.chatBtn);
                    chatBtn.textContent = "Chat";
                    chatBtn.onclick = () =>{
                      window.open(`msteams://teams.microsoft.com/l/chat/0/0?users=${item.Email}&message=${this.properties.message}`, '_blank');
                    };

                    buttonDiv.appendChild(profileSection);
                    buttonDiv.appendChild(chatBtn);

                    buttonsContainer!.appendChild(buttonDiv); // Append the button to the buttons container
                    buttonsCreated++; // Increment the count of buttons created
                } else {
                    // console.log("No button creation for: ", item.UserName);
                }
            });
            if (buttonsCreated === 0) {
              const noDataMessage: HTMLDivElement = document.createElement('div');
              noDataMessage.classList.add(styles.innerContents);
              noDataMessage.textContent = 'There are no new joinees the last 30 days';
              console.log("No new joinees the last 30 days");
              buttonsContainer!.appendChild(noDataMessage);// Non-null assertion operator
            }
        } else {
            const noDataMessage: HTMLDivElement = document.createElement('div');
            noDataMessage.classList.add(styles.innerContents);
            noDataMessage.textContent = 'There are no new joinees the last 30 days';
            console.log("No new joinees the last 30 days");
            buttonsContainer!.appendChild(noDataMessage);// Non-null assertion operator
        }
    })
    .catch(error => {
        console.error("Error fetching user data: ", error);
    });
}

private adjustDateForTimeZone(dateString) {
  // Add your timezone adjustment logic here
  const timeZoneDifferenceHours = 5; // Adjust this based on your timezone
  const timeZoneDifferenceMinutes = 30;

  const date = new Date(dateString);
  date.setHours(date.getHours() + timeZoneDifferenceHours);
  date.setMinutes(date.getMinutes() + timeZoneDifferenceMinutes);

  return date;
}

private formatBirthday(birthdayMonthDay: string): string {
  const [month, day] = birthdayMonthDay.split('-');
  return `${day}-${month}`;
}

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Title For The Application",
                }),
                PropertyPaneTextField('message', {
                  placeholder: "Welcome to the company!",
                  label: "Message for welcome wishes",
                }),
                // PropertyPaneToggle('displayBirthday', {
                //   label: 'Display Birthday',
                //   onText: 'Yes',
                //   offText: 'No',
                // }),
              ]
            }
          ]
        }
      ]
    };
  }
}
