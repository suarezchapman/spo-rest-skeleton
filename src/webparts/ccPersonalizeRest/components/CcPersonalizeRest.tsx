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
      userDisplayName,
      JokeText
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
          <h3>Logic App API Results</h3>

          <table className={styles.tables}>
            <tr>
              <td className={styles.tables}>
                {JokeText}
              </td>
            </tr>
          </table>

          <hr />

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

          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
          </ul>
        </div>
      </section>
    );
  }
}
