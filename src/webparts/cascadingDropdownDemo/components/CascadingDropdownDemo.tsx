import * as React from 'react';
import styles from './CascadingDropdownDemo.module.scss';
import { ICascadingDropdownDemoProps } from './ICascadingDropdownDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class CascadingDropdownDemo extends React.Component<ICascadingDropdownDemoProps, {}> {
  public render(): React.ReactElement<ICascadingDropdownDemoProps> {
    const {
      description,
      hasTeamsContext,
      list,
      fields
    } = this.props;

    return (
      <section className={`${styles.cascadingDropdownDemo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>{description}</h2>
          {
            list ? <>
              <p>Selected List: {list && list}</p>
              <p>Selected Fields: {fields.length && fields.toString()}</p>
            </> : <strong>Please select list from property pane</strong>
          }
        </div>
      </section>
    );
  }
}
