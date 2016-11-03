/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';

interface ICreateIssueProps {
	dispatch?: any;
	name?: string;
	content?: string;
}

function mapStateToProps(state: any): ICreateIssueProps {
	return ({
		name: state.controlState.name,
		content: state.controlState.content,
	});
}

@connect(mapStateToProps)

export class Issues extends React.Component<ICreateIssueProps, {}> {
	
	public constructor() {
		super();
		this.state = { let name: string = "", let content: string = ""};
		this.handleChange = this.handleChange.bind(this);
		this.handleSubmit = this.handleSubmit.bind(this);
	}

	public handleChange(event) {
		this.setState({name: event.target.value});
		this.setState({content: event.target.value});
	}

	public handleSubmit(event) : void {
		alert("Created an issue!");
	}

	public render(): React.ReactElement<Provider> {
		return (
			<div>
				<input type="text" placeholder="Enter in contents of the issue." name={this.state.name} content={this.state.content} onChange={this.handleChange} />
				<button onClick={this.handleSubmit}>Create Issue</button>
			</div>
		);
	}
}