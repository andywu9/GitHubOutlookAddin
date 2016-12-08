/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider, connect } from 'react-redux';
import * as UIFabric from 'office-ui-fabric-react';

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
		this.handleChangeName = this.handleChangeName.bind(this);
		this.handleChangeContent = this.handleChangeContent.bind(this);
		this.handleSubmit = this.handleSubmit.bind(this);
	}

	public handleChangeName(event) {
		this.setState({name: event.target.value});
	}

	public handleChangeContent(event) {
		this.setState({content: event.target.value});
	}

	public handleSubmit(event) : void {
		var myRepo = Office.context.roamingSettings.get("GitHub Repository");
		var repoArr = myRepo.split('/');
		var owner = repoArr[3];
		var repo = repoArr[4];

		var http = require('http');
		var querystring = require('querystring');
		var data = querystring.stringify({
		    title: this.state.name,
		    body: this.state.content
		  });


		var options = {
		  host: 'www.api.github.com',
		  path: '/repos/' + owner + '/' + repo + '/issues',
		};

		var callback = function(response) {
		  var str = ''
		  response.on('data', function (chunk) {
		    str += chunk;
		  });

		  response.on('end', function () {
		    console.log(str);
		  });
		}

		var req = http.request(options, callback);
		//This is the data we are posting, it needs to be a string or a buffer
		req.write(data);
		req.end();
	}

	public render(): React.ReactElement<Provider> {
		return (
			<div>
				<div>
					<UIFabric.TextField label='Name of the issue' onChanged={ this.handleChangeName } />
				</div>
				<div>
					<UIFabric.TextField label='Contents of the issue' multiline onChanged={ this.handleChangeContent } />
				</div>
				<button onClick={this.handleSubmit}>Create Issue</button>
				<h3>{this.state.name}</h3>
				<h2>{this.state.content}</h2>
			</div>
		);
	}
}
