import { Log } from '@microsoft/sp-core-library';
import {
  BaseFieldCustomizer,
  type IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import styles from './PriorityFieldCustomizerFieldCustomizer.module.scss';
import { getPriorityClass, buildBadgeText } from './priorityUtils';

const LOG_SOURCE: string = 'PriorityFieldCustomizerFieldCustomizer';

export interface IPriorityFieldCustomizerFieldCustomizerProperties {
  /** Optional prefix shown before the value (e.g. the field name). */
  prefix?: string;
}

/**
 * Field Customizer that renders the cell value as a colored badge depending
 * on the priority text: "high" -> red, "low" -> green, anything else -> grey.
 *
 * Covers the book's chapter 14 "Types of Extensions" (Field Customizer).
 */
export default class PriorityFieldCustomizerFieldCustomizer
  extends BaseFieldCustomizer<IPriorityFieldCustomizerFieldCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Activated PriorityFieldCustomizer');
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const value: string = (event.fieldValue as string) ?? '';
    const cls: string = getPriorityClass(value) === 'high'
      ? styles.high
      : getPriorityClass(value) === 'low'
        ? styles.low
        : styles.normal;

    event.domElement.classList.add(styles.priorityFieldCustomizer);
    event.domElement.innerText = '';

    const badge: HTMLSpanElement = document.createElement('span');
    badge.className = `${styles.badge} ${cls}`;
    badge.textContent = buildBadgeText(value, this.properties.prefix);
    event.domElement.appendChild(badge);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    super.onDisposeCell(event);
  }
}