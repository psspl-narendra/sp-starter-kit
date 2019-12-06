import * as React from 'react';
import styles from './Banner.module.scss';
import { IBannerProps } from '.';
import { escape } from '@microsoft/sp-lodash-subset';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import * as strings from 'BannerWebPartStrings';

export class Banner extends React.Component<IBannerProps, {}> {
  private _scrollElm: HTMLElement = null;
  private _scrollElmRect: ClientRect = null;
  private _parallaxElm: HTMLElement = null;

  /**
   * Opens the property pane of the web part for configuration
   */
  private _onConfigure = () => {
    this.props.propertyPane.open();
  }

  /**
   * Finds the scrollable parent
   */
  private _getScrollableParent(): HTMLElement {
    const scrollElm = document.querySelector('div[data-is-scrollable="true"]') as HTMLElement;
    if (scrollElm) {
      return scrollElm;
    }
    return null;
  }

  /**
   * Set translate / parallax effect
   * @param vector
   */
  private _setTranslate(vector: number) {
    const r = `translate3d(0px, ${vector}px, 0px)`;
    this._parallaxElm.style.transform = r;
  }

  /**
   * Set the parallax effect to the banner image element
   */
  private _setParallaxEffect = () => {
    window.requestAnimationFrame(() => {
      const scrollElmTop = this._scrollElmRect.top;
      const clientElmRect = this.props.domElement.getBoundingClientRect();
      const clientElmTop = clientElmRect.top;
      const clientElmBottom = clientElmRect.bottom;

      if (clientElmTop < scrollElmTop && clientElmBottom > scrollElmTop) {
        const vector = Math.round((scrollElmTop - clientElmTop) / 1.81);
        this._setTranslate(vector);
      } else if (clientElmTop >= scrollElmTop) {
        this._setTranslate(0);
      }
    });
  }

  private _removeParallaxBinding() {
    if (this._scrollElm) {
      // Unbind the scroll event
      this._scrollElm.removeEventListener("scroll", this._setParallaxEffect);
    }
  }

  /**
   * Add or remove parallax binding to the banner
   */
  private _parallaxBinding() {
    if (this.props.useParallaxInt) {
      this._scrollElm = this._getScrollableParent();
      this._parallaxElm = this.props.domElement.querySelector(`.${styles.bannerImg}`) as HTMLElement;
      if (this._scrollElm && this._parallaxElm) {
        // Get client rect info
        this._scrollElmRect = this._scrollElm.getBoundingClientRect();
        // Bind the scroll event
        this._scrollElm.addEventListener("scroll", this._setParallaxEffect);
      }
    } else {
      this._removeParallaxBinding();
    }
  }

  /**
   * Bind the parallax effect on mount
   */
  public componentDidMount(): void {
    this._parallaxBinding();
  }

  /**
   * Bind the parallax effect on update
   * @param prevProps
   */
  public componentDidUpdate(prevProps: IBannerProps): void {
    this._parallaxBinding();
  }

  /**
   * Unbind the parallax effect
   */
  public componentWillUnmount(): void {
    this._removeParallaxBinding();
  }

  /**
   * Default React render method
   */
  public render(): React.ReactElement<IBannerProps> {
    if (this.props.bannerImage) {
      return (
        <div className={styles.banner} style={{
          height: this.props.bannerHeight ? `${this.props.bannerHeight}px` : `280px`
        }}>
          <div className={styles.bannerImg} style={{
            backgroundImage: `url('${this.props.bannerImage}')`
          }}></div>
          <div className={styles.bannerOverlay}></div>
          <div className={styles.bannerText}>
            {
              this.props.bannerLink ? (
                <a href={this.props.bannerLink} title={escape(this.props.bannerText)}>{escape(this.props.bannerText)}</a>
              ) : (
                <span>{escape(this.props.bannerText)}</span>
              )
            }
          </div>
        </div>
      );
    } else {
      return <Placeholder iconName='ImagePixel'
                          iconText={strings.BannerPlaceholderIconText}
                          description={strings.BannerPlaceholderDescription}
                          buttonLabel={strings.BannerPlaceholderBtnLabel}
                          onConfigure={this._onConfigure} />;
    }
  }
}
