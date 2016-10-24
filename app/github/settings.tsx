/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';

export class Settings extends React.Component< any, any >{

	public constructor(props: any) {
		super(props);
		this.state = { let isFirsttime: boolean = true, let repo: string = "" };
		this.handleChange = this.handleChange.bind(this);
		this.handleSubmit = this.handleSubmit.bind(this);
	}

	public handleChange(event) {
		this.setState({isFirsttime: false});
		this.setState({repo: event.target.value});
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