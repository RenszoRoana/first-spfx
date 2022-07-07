var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
var HelloWorldWebPart = /** @class */ (function (_super) {
    __extends(HelloWorldWebPart, _super);
    function HelloWorldWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._environmentMessage = '';
        return _this;
    }
    HelloWorldWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n    <section class=\"".concat(styles.helloWorld, " ").concat(!!this.context.sdks.microsoftTeams ? styles.teams : '', "\">\n      <div class=\"").concat(styles.welcome, "\">\n        <img alt=\"\" src=\"").concat(this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png'), "\" class=\"").concat(styles.welcomeImage, "\" />\n        <h2>Well done, ").concat(escape(this.context.pageContext.user.displayName), "!</h2>\n        <div>").concat(this._environmentMessage, "</div>\n        <div>Web part property value: <strong>").concat(escape(this.properties.description), "</strong></div>\n      </div>\n      <div>\n        <h3>Welcome to SharePoint Framework!</h3>\n        <p>\n        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.\n        </p>\n        <h4>Learn more about SPFx development:</h4>\n          <ul class=\"").concat(styles.links, "\">\n            <li><a href=\"https://aka.ms/spfx\" target=\"_blank\">SharePoint Framework Overview</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-graph\" target=\"_blank\">Use Microsoft Graph in your solution</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-teams\" target=\"_blank\">Build for Microsoft Teams using SharePoint Framework</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-viva\" target=\"_blank\">Build for Microsoft Viva Connections using SharePoint Framework</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-store\" target=\"_blank\">Publish SharePoint Framework applications to the marketplace</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-api\" target=\"_blank\">SharePoint Framework API reference</a></li>\n            <li><a href=\"https://aka.ms/m365pnp\" target=\"_blank\">Microsoft 365 Developer Community</a></li>\n          </ul>\n      </div>\n    </section>");
    };
    HelloWorldWebPart.prototype.onInit = function () {
        this._environmentMessage = this._getEnvironmentMessage();
        return _super.prototype.onInit.call(this);
    };
    HelloWorldWebPart.prototype._getEnvironmentMessage = function () {
        if (!!this.context.sdks.microsoftTeams) { // running in Teams
            return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
        }
        return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
    };
    HelloWorldWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }
    };
    Object.defineProperty(HelloWorldWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    HelloWorldWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return HelloWorldWebPart;
}(BaseClientSideWebPart));
export default HelloWorldWebPart;
//# sourceMappingURL=HelloWorldWebPart.js.map