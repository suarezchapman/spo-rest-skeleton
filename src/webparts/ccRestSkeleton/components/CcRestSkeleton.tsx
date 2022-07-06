import * as React from 'react';
import styles from './CcRestSkeleton.module.scss';
import { ICcRestSkeletonProps } from './ICcRestSkeletonProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class CcRestSkeleton extends React.Component<ICcRestSkeletonProps, {}> {

  public render(): React.ReactElement<ICcRestSkeletonProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userLoginName,
      userEmail,
      userDisplayName,
      APIResult
    } = this.props;

    const LoginName = this.props.userLoginName;
    const LoginNameArray = LoginName.split("@");
    const UserName = LoginNameArray[0];

    const data = JSON.parse(this.props.APIResult);

    const MDDUserID = data[0]["MDDUserID"];
    const MDDLocationCubeID = data[0]["MDDLocationCubeID"];
    const PhoneOffice = data[0]["PhoneOffice"];
    const ccDrupalPrimaryGroup = data[0]["ccDrupalPrimaryGroup"];

    console.log(data[0]["HomeAddressStreet1"]);

    return (
      <section className={`${styles.ccRestSkeleton} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : 'https://webapps.chapman.com/fromneverest/ea-html/930/93/09/images/' + UserName + '.jpg'} className={styles.welcomeImage} />
          <h2>{escape(userDisplayName)}</h2>
          <div>UserName:  <strong>{UserName}</strong></div>
          <div>environmentMessage:  <strong>{environmentMessage}</strong></div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Page Properties (pageContext.user)</h3>

          <table className={styles.tables}>
            <tr>
              <td className={styles.tables}>
                userDiplayName
              </td>
              <td className={styles.tables}>
                {escape(userDisplayName)}
              </td>
            </tr>
            <tr>
              <td className={styles.tables}>
                <strong>userLoginName</strong>
              </td>
              <td className={styles.tables}>
                <strong>{escape(userLoginName)}</strong>
              </td>
            </tr>
            <tr>
              <td className={styles.tables}>
                userEmail
              </td>
              <td className={styles.tables}>
                {escape(userEmail)}
              </td>
            </tr>
          </table>

          <hr />

          <h3>Logic App API Results from MDD Based on Login User</h3>

          <table className={styles.tables}>
            <tr>
              <td className={styles.tables}>
                {APIResult}
              </td>
            </tr>
          </table>

          <hr />

          <h3>MDD Data from Above Personalized JSON Result</h3>

          <table className={styles.tables}>
            <tr>
              <td className={styles.tables}>
                MDDUserID
              </td>
              <td className={styles.tables}>
                {escape(MDDUserID)}
              </td>
            </tr>
            <tr>
              <td className={styles.tables}>
                MDDLocationCubeID
              </td>
              <td className={styles.tables}>
                {escape(MDDLocationCubeID)}
              </td>
            </tr>
            <tr>
              <td className={styles.tables}>
                PhoneOffice
              </td>
              <td className={styles.tables}>
                {escape(PhoneOffice)}
              </td>
            </tr>
            <tr>
              <td className={styles.tables}>
                ccDrupalPrimaryGroup
              </td>
              <td className={styles.tables}>
                {escape(ccDrupalPrimaryGroup)}
              </td>
            </tr>
          </table>
        </div>
      </section>
    );
  }
}
