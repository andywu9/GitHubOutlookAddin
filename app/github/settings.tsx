/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import * as UIFabric from 'office-ui-fabric-react';

interface IChangeRepositoryProps {
	dispatch?: any;
	firstTimeState?: boolean;
	repo?: string;
}

function mapStateToProps(state: any): IChangeRepositoryProps {
	return ({
		firstTimeState: state.controlState.firstTimeState,
		repo: state.controlState.repo,
	});
}

@connect(mapStateToProps)

export class Settings extends React.Component<IChangeRepositoryProps, any> {

	public constructor() {
		super();
		this.state = { firstTimeState: true, repo: "" };
		this.handleChange = this.handleChange.bind(this);
		this.handleSubmit = this.handleSubmit.bind(this);
	}

	public handleChange(event) {
		this.setState({firstTimeState: false});
		this.setState({repo: event.target.value});
		console.log(this.state.repo);
	}

	public handleSubmit(event) : void {
		Office.context.roamingSettings.set("GitHub Repository", this.state.repo);
		Office.context.roamingSettings.saveAsync();
	}

	public render(): React.ReactElement<Provider> {
		return (
			<div>
				<UIFabric.TextField label='Github Repository' onChanged={ this.handleChange } />
				<div>
					<UIFabric.Button onClick={this.handleSubmit}>Change repository</UIFabric.Button>
				</div>
			</div>
		);
	}
}
