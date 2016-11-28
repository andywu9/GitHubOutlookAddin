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

export class Issues extends React.Component<ICreateIssueProps, any> {

	public constructor() {
		super();
		this.state = { name: "", content: "" };
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
				<div>
					<h2>Name of the issue</h2>
					<input type="text" placeholder="Enter the name of the issue." name={this.state.name} content={this.state.content} onChange={this.handleChange} />
				</div>
				<div> 
					<h2>Contents of the issue</h2>
					<input type="text" placeholder="Enter in contents of the issue." name={this.state.name} content={this.state.content} onChange={this.handleChange} />
				</div>
				<button onClick={this.handleSubmit}>Create Issue</button>
				<h3>{this.state.name}</h3>
				<h2>{this.state.content}</h2>
			</div>
		);
	}
}
