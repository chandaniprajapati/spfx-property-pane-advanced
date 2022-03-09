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
      filePickerResult
    } = this.props;

    return (
      <section className={`${styles.dynamicPropertyPane} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
          <h2>{description}</h2>
          {this.props.textOrImageType == "Text" ? this.props.simpleText :
            (filePickerResult && <img src={filePickerResult.fileAbsoluteUrl} height="100px" width="100px" alt={filePickerResult.fileName} />)
          }
          {/* {textOrImageType == "Text" ? simpleText : filePickerResult && filePickerResult.fileName} */}
        </div>
      </section>
    );
  }
}
