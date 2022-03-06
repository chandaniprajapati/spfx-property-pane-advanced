import * as React from 'react';
import styles from './DynamicPropertyPane.module.scss';
import { IDynamicPropertyPaneProps } from './IDynamicPropertyPaneProps';

export default class DynamicPropertyPane extends React.Component<IDynamicPropertyPaneProps, {}> {
  public render(): React.ReactElement<IDynamicPropertyPaneProps> {
    const {
      hasTeamsContext,
      textOrImageType,
      description,
      simpleText,
      imageUrl
    } = this.props;

    return (
      <section className={`${styles.dynamicPropertyPane} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
          <h2>{description}</h2>
          {textOrImageType == "Text" ? simpleText :
            (<img src={imageUrl} height="250px" width="250px" alt="No Image" />)
          }
        </div>
      </section>
    );
  }
}
