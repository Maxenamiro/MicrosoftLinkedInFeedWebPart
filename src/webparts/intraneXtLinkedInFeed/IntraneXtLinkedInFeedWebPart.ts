import { Version } from '@microsoft/sp-core-library'
import {
	IPropertyPaneConfiguration,
	PropertyPaneTextField,
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import styles from './IntraneXtLinkedInFeedWebPart.module.scss'

export interface IIntraneXtLinkedInFeedWebPartProps {
	elfSiteUrl: string
}
//
export default class IntraneXtLinkedInFeedWebPart extends BaseClientSideWebPart<IIntraneXtLinkedInFeedWebPartProps> {
	public render(): void {
		const { elfSiteUrl } = this.properties

		if (!elfSiteUrl) {
			this.domElement.innerHTML = `
        <div class="${styles.infoFeed}">
          <div class="${styles.placeholder}">
            Please configure the Elf Site URL in the property pane.
          </div>
        </div>`
			return
		}

		const safeUrl =
			elfSiteUrl.indexOf('http') === 0 ? elfSiteUrl : `https://${elfSiteUrl}`

		this.domElement.innerHTML = `
      <div class="${styles.infoFeed}">
        <iframe
          src="${safeUrl}"
          class="${styles.iframe}"
          frameborder="0"
          sandbox="allow-same-origin allow-scripts allow-forms allow-popups"
          referrerpolicy="no-referrer"
        ></iframe>
      </div>`
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0')
	}

	protected onPropertyPaneFieldChanged(
		propertyPath: string,
		oldValue: any,
		newValue: any
	): void {
		if (oldValue !== newValue) {
			this.render()
		}
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: { description: 'Elf Feed Viewer Settings' },
					groups: [
						{
							groupName: 'Feed Configuration',
							groupFields: [
								PropertyPaneTextField('elfSiteUrl', {
									label: 'Elf Site URL',
									placeholder: 'https://yoursite.elf.site/',
									description: 'Enter the full URL of your Elf site',
								}),
							],
						},
					],
				},
			],
		}
	}
}
