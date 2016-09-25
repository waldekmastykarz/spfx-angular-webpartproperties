import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import styles from './HelloWorld.module.scss';
import * as strings from 'helloWorldStrings';
import { IHelloWorldWebPartProps } from './IHelloWorldWebPartProps';
import * as angular from 'angular';

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private $injector: ng.auto.IInjectorService;

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    console.log('render');
    if (!this.renderedOnce) {
      this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}" ng-controller="HomeController as vm">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">{{vm.hello}}</p>
              <a href="#" class="ms-Button ${styles.button}" ng-click="vm.click($event)">
                <span class="ms-Button-label">ng click</span>
              </a>
            </div>
          </div>
        </div>
        <div class="${styles.container}" ng-controller="SecondController as vm">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">{{vm.hello}}</p>
            </div>
          </div>
        </div>
      </div>`;

      angular.module('helloWorld', [])
        .controller('HomeController', ['$rootScope', '$scope', function ($rootScope: ng.IRootScopeService, $scope: ng.IScope) {
          $rootScope.$on('propertiesChanged', (event: ng.IAngularEvent, properties: IHelloWorldWebPartProps): void => {
            this.hello = properties.description;
          });

          (this as any).click = ($event: ng.IAngularEvent) => {
            $event.preventDefault();
            this.hello = 'Modified from Angular';
            $rootScope.$broadcast('propertyChanged', 'description', this.hello);
          }
        }])
        .controller('SecondController', ['$rootScope', '$scope', function ($rootScope: ng.IRootScopeService, $scope: ng.IScope) {
          $rootScope.$on('propertiesChanged', (event: ng.IAngularEvent, properties: IHelloWorldWebPartProps): void => {
            this.hello = properties.description;
          });
        }]);

      this.$injector = angular.bootstrap(this.domElement, ['helloWorld']);
      this.$injector.get('$rootScope').$on('propertyChanged', (event: ng.IAngularEvent, propertyName, propertyValue): void => {
        this.onPropertyChange(propertyName, propertyValue);
      });
    }

    this.$injector.get('$rootScope').$broadcast('propertiesChanged', this.properties);
    if (!this.renderedOnce) {
      this.$injector.get('$rootScope').$digest();
    }
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
  }
}
