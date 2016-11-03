/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';

interface IChangeRepositoryProps {
	dispatch?: any;
	firstTimeState?: boolean;
	repo: string;
}

function mapStateToProps(state: any): IChangeRepositoryProps {
	return ({
	firstTimeState: state.controlState.firstTimeState,
	repo: state.controlState.repo,
	});
}

@connect(mapStateToProps)

export class Settings extends React.Component<IChangeRepositoryProps, {}> {

	public constructor() {
		super();
		this.state = { let firstTimeState: boolean = true, repo: string = "" };
		this.handleChange = this.handleChange.bind(this);
		this.handleSubmit = this.handleSubmit.bind(this);
	}

	public handleChange(event) {
		this.setState({firstTimeState: false});
		this.setState({repo: event.target.value});
		Office.context.roamingSettings.set("GitHub Repository", this.state.repo);
		Office.context.roamingSettings.saveAsync();
	}

	public handleSubmit(event) : void {
		alert("Changed repository!");
	}

	public render(): React.ReactElement<Provider> {
		return (
			<input type="text" placeholder="Enter in the link to the Github repository." repo={this.state.repo} onChange={this.handleChange} />
				<button onClick={this.handleSubmit}>Change repository</button>
		);
	}
}