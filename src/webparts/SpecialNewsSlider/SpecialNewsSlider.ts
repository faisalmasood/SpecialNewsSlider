import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IWebPartContext,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider
} from "@microsoft/sp-webpart-base";
import { Version, DisplayMode } from "@microsoft/sp-core-library";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import styles from "./SpecialNewsSlider.module.scss";
import * as strings from "SpecialNewsSliderStrings";
import { ISpecialNewsSliderProps } from "./ISpecialNewsSliderProps";
import * as React from "react";
import * as ReactDom from "react-dom";
// Imports property pane custom fields
// Don't change the PropertyFieldCustomList, CustomListFieldType from official sp-client-custom-fields path except you're sure PropertyFieldCustomList component issue was fixed.
import {
  PropertyFieldCustomList,
  CustomListFieldType
} from "./components/PropertyFieldCustomList";
import { PropertyFieldColorPickerMini } from "sp-client-custom-fields/lib/PropertyFieldColorPickerMini";
import { PropertyFieldFontPicker } from "sp-client-custom-fields/lib/PropertyFieldFontPicker";
import { PropertyFieldFontSizePicker } from "sp-client-custom-fields/lib/PropertyFieldFontSizePicker";
import { PropertyFieldAlignPicker } from "sp-client-custom-fields/lib/PropertyFieldAlignPicker";
import SpecialNewsSliderScriptLoader from "./SpecialNewsSliderScriptLoader";

// from https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/use-sp-pnp-js-with-spfx-web-parts
// SPFX
import { sp, CamlQuery, Web } from "@pnp/sp";

// Loads external CSS files
require("./unitegallery/styles/unite-gallery.scss");

/**
 * @interface
 * The interface is used to store web part title react component properties.
 */
export interface IWebPartTitle {
  title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}

/**
 * @interface
 * Used to store the scoop data.
 */
export interface IScoop {
  Enable: boolean;
  Title: string;
  ListTitle: string;
  ViewTitle: string;
}

/**
 * @class
 * The Scoop web part class.
 */
export default class SpecialNewsSlider extends BaseClientSideWebPart<
  ISpecialNewsSliderProps
> {
  // SPFX Init
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  // The private field that is used to store guid value.
  private guid: string;
  /**
   * @constructor
   * The constructor method for web part.
   * @param context The web part context object.
   */
  public constructor(context?: IWebPartContext) {
    super();
    this.guid = this.getGuid();
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(
      this
    );
  }

  /**
   * @function
   * Entry render method which is used to render web part HTML elements.
   */
  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.SpecialNewsSlider}">
      <div class="${styles.SpecialNewsSliderTitle}"></div>
      <div class="${styles.SpecialNewsSliderMessageContainer}"></div>
      <div class="${styles.SpecialNewsSliderGallery}"></div>
    </div>
    `;
    this.renderWebpartTitle();
    this.getsources().then(
      (sources: IScoop[]) => {
        this.renderScoopGallery(sources);
      },
      error => {
        console.error(error);
      }
    );
  }

  /**
   * @function
   * Render the web part title.
   */
  private renderWebpartTitle() {
    const element: React.ReactElement<IWebPartTitle> = React.createElement(
      WebPartTitle,
      {
        title: this.properties.webPartTitle,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.webPartTitle = value;
        }
      }
    );
    ReactDom.render(
      element,
      this.domElement.getElementsByClassName(styles.SpecialNewsSliderTitle)[0]
    );
  }

  /**
   * @function
   * Render the data gallery.
   * @param sources The scoop data objects.
   */
  private renderScoopGallery(sources: IScoop[]) {
    if (
      sources === null ||
      sources === undefined ||
      (sources !== null && sources !== undefined && sources.length == 0)
    ) {
      this.renderMessage(strings.SpecialNewsSliderNoDataMessage);
    } else {
      let jqueryScript: SpecialNewsSliderScriptLoader.IScript = {
        Url: "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.11.1.min.js",
        GlobalExportsName: "jQuery",
        ExtendedTarget: window,
        ExtendedPropertyOrFunctionChain: "jQuery"
      };
      let dependencies: SpecialNewsSliderScriptLoader.IScript[] = [];
      SpecialNewsSliderScriptLoader.LoadScript(jqueryScript, dependencies).then(
        object => {
          require("./unitegallery/scripts/unitegallery.js");
          SpecialNewsSliderScriptLoader.WaitExtendedObjectAttached(
            ($ as any)("#" + this.guid + "-gallery"),
            "unitegallery"
          ).then(
            IsLoaded => {
              if (IsLoaded) {
                // Render HTML
                let results = [];
                var galleryHtml: string = `<div id="${
                  this.guid
                }-gallery" style="display:none;">`;

                // Promise Array
                let getSPListItems = [];
                for (var i = 0; i < sources.length; i++) {
                  // Skip
                  var scoop: IScoop = sources[i];
                  if (scoop.Enable === false) continue;

                  // Current Web context
                  var webpart = this;
                  let currentWeb: Web = new Web(scoop.Title);

                  // SPFX
                  // from https://github.com/SharePoint/PnP-js-core/wiki
                  // Loop SPList GET query
                  // from https://sharepoint.stackexchange.com/questions/135936/how-to-get-all-items-in-a-view-using-rest-api
                  
                  // Collect promise array
                  getSPListItems.push(
                    currentWeb.lists
                      .getByTitle(scoop.ListTitle)
                      .get()
                      .then(function(listResp) {
                        return currentWeb.lists
                          .getByTitle(scoop.ListTitle)
                          .views
                          .getByTitle(scoop.ViewTitle)
                          .get()
                          .then(function(viewResp) {
                            // SFPX
                            // from https://github.com/SharePoint/PnP-JS-Core/wiki/Working-With:-Items
                            return currentWeb.lists
                              .getByTitle(listResp.Title)
                              .getItemsByCAMLQuery({
                                ViewXml: viewResp.ListViewXml
                              })
                              .then(function(itemResp) {
                                return itemResp;
                              });
                          });
                      })
                  );
                }

                // After REST api
                Promise.all(getSPListItems).then(function(values) {
                  // Collect items
                  console.log("Render");
                  // Outer loop - SPList
                  for (var i = 0; i < values.length; i++) {
                    // Inner loop - SPListIem
                    for (var j = 0; j < values[i].length; j++) {
                      // Append
                      let result = values[i][j];
                      galleryHtml += `<a href="${result.LinkUrl}">
                                      <img alt="${result.Title}" src="${
                        result.LinkUrl
                      }" data-image="${result.Picture}" data-description="${
                        result.Description
                      }"/></a>`;
                    }
                  }

                  // Render HTML
                  galleryHtml += "</div>";
                  console.log(galleryHtml);

                  // Initialize JQuery
                  $(webpart.domElement)
                    .find("." + styles.SpecialNewsSliderGallery)
                    .html(galleryHtml);
                  webpart.unitegallery();
                });
              }
            },
            error => {
              console.error(error);
            }
          );
        },
        error => {
          console.error(error);
        }
      );
    }
  }

  /**
   * @function
   * Calling unitegallery third-party library to render the data.
   */
  private unitegallery(): void {
    try {
      ($ as any)("#" + this.guid + "-gallery").unitegallery({
        gallery_theme: "slider",
        slider_enable_arrows: this.properties.enableArrows,
        slider_enable_bullets: this.properties.enableBullets,
        slider_transition: this.properties.transition,
        gallery_preserve_ratio: this.properties.preserveRatio,
        gallery_autoplay: this.properties.autoplay,
        gallery_play_interval: this.properties.speed,
        gallery_pause_on_mouseover: this.properties.pauseOnMouseover,
        gallery_carousel: this.properties.carousel,
        slider_enable_progress_indicator: this.properties
          .enableProgressIndicator,
        slider_enable_play_button: this.properties.enablePlayButton,
        slider_enable_fullscreen_button: this.properties.enableFullscreenButton,
        slider_enable_zoom_panel: this.properties.enableZoomPanel,
        slider_controls_always_on: this.properties.controlsAlwaysOn,
        slider_enable_text_panel: this.properties.textPanelEnable,
        slider_textpanel_always_on: this.properties.textPanelAlwaysOnTop,
        slider_textpanel_bg_color: this.properties.textPanelBackgroundColor,
        slider_textpanel_bg_opacity: this.properties.textPanelOpacity,
        slider_textpanel_title_color: this.properties.textPanelFontColor,
        slider_textpanel_title_font_family: this.properties.textPanelFont,
        slider_textpanel_title_text_align: this.properties.textPanelAlign,
        slider_textpanel_title_font_size:
          this.properties.textPanelFontSize != null
            ? this.properties.textPanelFontSize.replace("px", "")
            : ""
      });
    } catch (error) {
      console.error(error);
    }
  }

  /**
   * @function
   * Render the messsage.
   * @param message The message to be rendered.
   */
  private renderMessage(message: string) {
    let messageHtml = `
      <div class="ms-MessageBar">
        <div class="ms-MessageBar-content">
          <div class="ms-MessageBar-icon">
            <i class="ms-Icon ms-Icon--Info"></i>
          </div>
          <div class="ms-MessageBar-text">${message}</div>
        </div>
      </div>`;
    this.domElement.querySelector(
      "." + styles.SpecialNewsSliderMessageContainer
    ).innerHTML = messageHtml;
  }

  /**
   * @function
   * Generate globally unique identifier.
   */
  private getGuid(): string {
    return (
      this.s4() +
      this.s4() +
      "-" +
      this.s4() +
      "-" +
      this.s4() +
      "-" +
      this.s4() +
      "-" +
      this.s4() +
      this.s4() +
      this.s4()
    );
  }

  /**
   * @function
   * Generate 4 random characters.
   */
  private s4(): string {
    return Math.floor((1 + Math.random()) * 0x10000)
      .toString(16)
      .substring(1);
  }

  /**
   * @property
   * Set the data version.
   */
  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  /**
   * @function
   * This function is used to configure the web part's configuration panel.
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.SpecialNewsSliderPropertyPageGeneralPanel
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.SpecialNewsSliderGeneralPanelDataGroupName,
              groupFields: [
                PropertyFieldCustomList("sources", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelDataManagementLabel,
                  value: this.properties.sources,
                  headerText:
                    strings.SpecialNewsSliderGeneralPanelDataManagementTableHeaderText,
                  fields: [
                    {
                      id: "Title",
                      title: "Web URL",
                      required: true,
                      type: CustomListFieldType.string
                    },
                    {
                      id: "Enable",
                      title:
                        strings.SpecialNewsSliderGeneralPanelDataManagementTableEnableColumnDisplayName,
                      required: true,
                      type: CustomListFieldType.boolean
                    },
                    {
                      id: "ListTitle",
                      title: "List",
                      required: false,
                      hidden: true,
                      type: CustomListFieldType.string
                    },
                    {
                      id: "ViewTitle",
                      title: "View",
                      required: true,
                      hidden: true,
                      type: CustomListFieldType.string
                    }
                  ],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this
                    .disableReactivePropertyChanges,
                  context: this.context,
                  properties: this.properties,
                  key: "SpecialNewsSliderListField"
                })
              ]
            },
            {
              groupName: strings.SpecialNewsSliderGeneralPanelGroupName,
              groupFields: [
                PropertyPaneToggle("enableArrows", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelEnableArrowsToggleLabel
                }),
                PropertyPaneToggle("enableBullets", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelEnableBulletsToggleLabel
                }),
                PropertyPaneToggle("enableProgressIndicator", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelEnableProgressIndicatorToggleLabel
                }),
                PropertyPaneToggle("enablePlayButton", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelEnablePlayToggleLabel
                }),
                PropertyPaneToggle("enableFullscreenButton", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelEnableFullscreenToggleLabel
                }),
                PropertyPaneToggle("enableZoomPanel", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelEnableZoomPanelToggleLabel
                }),
                PropertyPaneToggle("controlsAlwaysOn", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelControlsAlwaysOnToggleLabel
                })
              ]
            },
            {
              groupName: strings.SpecialNewsSliderGeneralPanelEffectsGroupName,
              groupFields: [
                PropertyPaneDropdown("transition", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelTransitionDropdownLabel,
                  options: [
                    {
                      key: "slide",
                      text:
                        strings.SpecialNewsSliderGeneralPanelTransitionDropdownSlideOptionText
                    },
                    {
                      key: "fade",
                      text:
                        strings.SpecialNewsSliderGeneralPanelTransitionDropdownFadeOptionText
                    }
                  ]
                }),
                PropertyPaneToggle("preserveRatio", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelPreserveRatioToggleLabel
                }),
                PropertyPaneToggle("pauseOnMouseover", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelPauseOnMouseoverToggleLabel
                }),
                PropertyPaneToggle("carousel", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelCarouselToggleLabel
                }),
                PropertyPaneToggle("autoplay", {
                  label:
                    strings.SpecialNewsSliderGeneralPanelAutoplayToggleLabel
                }),
                PropertyPaneSlider("speed", {
                  label: strings.SpecialNewsSliderGeneralPanelSpeedSliderLabel,
                  min: 0,
                  max: 7000,
                  step: 100
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.SpecialNewsSliderPropertyPageTextPanel
          },
          groups: [
            {
              groupName: strings.SpecialNewsSliderTextPanelGroupName,
              groupFields: [
                PropertyPaneToggle("textPanelEnable", {
                  label: strings.SpecialNewsSliderTextPanelEnableFieldLabel
                }),
                PropertyPaneToggle("textPanelAlwaysOnTop", {
                  label: strings.SpecialNewsSliderTextPanelAlwaysOnTopFieldLabel
                }),
                PropertyPaneSlider("textPanelOpacity", {
                  label: strings.SpecialNewsSliderTextPanelOpacityFieldLabel,
                  min: 0,
                  max: 1,
                  step: 0.1
                }),
                PropertyFieldAlignPicker("textPanelAlign", {
                  label: strings.SpecialNewsSliderTextPanelAlignFieldLabel,
                  initialValue: this.properties.textPanelAlign,
                  onPropertyChanged: this.onPropertyPaneFieldChanged,
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this
                    .disableReactivePropertyChanges,
                  properties: this.properties,
                  key: "SpecialNewsSliderAlignField"
                }),
                PropertyFieldFontPicker("textPanelFont", {
                  label: strings.SpecialNewsSliderTextPanelFontFieldLabel,
                  initialValue: this.properties.textPanelFont,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this
                    .disableReactivePropertyChanges,
                  properties: this.properties,
                  key: "SpecialNewsSliderFontField"
                }),
                PropertyFieldFontSizePicker("textPanelFontSize", {
                  label: strings.SpecialNewsSliderTextPanelFontSizeFieldLabel,
                  initialValue: this.properties.textPanelFontSize,
                  usePixels: true,
                  preview: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this
                    .disableReactivePropertyChanges,
                  properties: this.properties,
                  key: "SpecialNewsSliderFontSizeField"
                }),
                PropertyFieldColorPickerMini("textPanelFontColor", {
                  label: strings.SpecialNewsSliderTextPanelFontColorFieldLabel,
                  initialColor: this.properties.textPanelFontColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this
                    .disableReactivePropertyChanges,
                  properties: this.properties,
                  key: "SpecialNewsSliderFontColorField"
                }),
                PropertyFieldColorPickerMini("textPanelBackgroundColor", {
                  label:
                    strings.SpecialNewsSliderTextPanelBackgroundColorFieldLabel,
                  initialColor: this.properties.textPanelBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this
                    .disableReactivePropertyChanges,
                  properties: this.properties,
                  key: "SpecialNewsSliderBgColorField"
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * @property
   * Set whether use reactive mode or not for property panel when property changed.
   */
  protected get disableReactivePropertyChanges(): boolean {
    // We do not want to disable reactive property changes, because the third-party custom fields used by this solution require reactive property changes to be enabled.
    return false;
  }

  /**
   * @function
   * Get data from different data sources.
   */
  protected getsources(): Promise<IScoop[]> {
    return new Promise<IScoop[]>((resolve, reject) => {
      let sources: IScoop[] = [];
      let webPartsources = this.getsourcesFromWebPart();
      // As we plan to add additional data sources in the future, We will need to call the functions to get data from other data sources here.
      sources = sources.concat(webPartsources);
      resolve(sources);
    });
  }

  /**
   * @function
   * Get data from the web part itself.
   */
  protected getsourcesFromWebPart(): IScoop[] {
    return this.properties.sources.map(scoopItem => {
      return {
        Title: scoopItem["Title"],
        Enable: scoopItem["Enable"].toString() == "true" ? true : false,
        ListTitle: scoopItem["ListTitle"],
        ViewTitle: scoopItem["ViewTitle"]
      };
    });
  }
}
