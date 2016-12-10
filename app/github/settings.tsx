/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import * as UIFabric from 'office-ui-fabric-react';

/**
 *  Properties needed for the Settings component
 *  @interface IChangeRepository Props
 */
interface IChangeRepositoryProps {
	repo?: string;
}

/**
 * maps state in application store to properties for the component
 * @param {any} state
 */
function mapStateToProps(state: any): IChangeRepositoryProps {
	return ({
		repo: state.controlState.repo,
	});
}

@connect(mapStateToProps)

export class Settings extends React.Component<IChangeRepositoryProps, any> {

	public constructor() {
		super();
		this.state = { repo: "" };
		this.handleChange = this.handleChange.bind(this);
		this.handleSubmit = this.handleSubmit.bind(this);
	}

	/**
	 * updates the state to reflect changes in the "GitHub Repository" TextField
	 */
	public handleChange(text) {
		console.log("Previous repo: " + this.state.repo);
		this.setState({repo: text});
		console.log("New repo: " + this.state.repo);
	}

	/**
	 * Changes and persists the repository in Outlook settings when "Change Repository" is clicked
	 */
	public handleSubmit(event) : void {
		//Office.context.roamingSettings.set("GitHub Repository", this.state.repo);
		//Office.context.roamingSettings.saveAsync();
		//var myRepo = Office.context.roamingSettings.get("GitHub Repository");
		//console.log("This is my saved repo: " + myRepo);
		console.log("Clicked change repo to: " + this.state.repo);
	}

	/**
	 * Renders the form
	 */
	public render(): React.ReactElement<Provider> {
		return (
			<div>
				<UIFabric.TextField label='Github Repository' onChanged={this.handleChange} />

				<div>
					<UIFabric.Button onClick={this.handleSubmit}>Change Repository</UIFabric.Button>
				</div>
			</div>
		);
	}
}
