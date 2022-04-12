import * as React from 'react';
import styles from './CcPersonalizeRest.module.scss';
import { ICcPersonalizeRestProps } from './ICcPersonalizeRestProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class CcPersonalizeRest extends React.Component<ICcPersonalizeRestProps, {}> {







  public render(): React.ReactElement<ICcPersonalizeRestProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userLoginName,
      userEmail,
      userDisplayName
    } = this.props;




    
    return (
      <section className={`${styles.ccPersonalizeRest} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
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
              userLoginName
              </td>
              <td className={styles.tables}>
              {escape(userLoginName)}
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
          <ul>
            <li>userDiplayName: {escape(userDisplayName)}</li>
            <li>userLoginName: {escape(userLoginName)}</li>
            <li>userEmail: {escape(userEmail)}</li>
          </ul>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }
}
