import * as React from 'react';
import { ICountriesViewProps } from '../../models/CountriesModel';
import './Countries.scss';
import SaveButton from '../../shared/components/SaveButton/SaveButton';
import CancelButton from '../../shared/components/CancelButton/CancelButton';
import PopUpWindow from '../../shared/components/PopUpWindow/PopUpWindow';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import Row from '../../shared/components/SingleRow/SingleRow';

export default class CountriesFormView extends React.Component<ICountriesViewProps> {
	public render(): React.ReactElement<ICountriesViewProps> {
		return (
			<div className="formWrapper">
				{this.props.isSubmitting && (
					<div className="spinnerOverlay">
						<Spinner size={SpinnerSize.large} label="Submitting contract..." />
					</div>
				)}
				{this.props.apiMessage && (
					<PopUpWindow
						onClose={this.props.closePopUpWindow}
						success={this.props.apiMessage.success}
						message={this.props.apiMessage.message}
						closeButton={this.props.PopUpWindowCloseButton}
					/>
				)}
				<div className="titleRow">
					<h1 className='title'>Countries</h1>
				</div>
				<form onSubmit={this.props.onSubmit} className="formContainer">
									<div className="buttonRow">
						<SaveButton />
						<CancelButton onClick={this.props.onCancel} />
					</div>
				</form>
			</div>
		);
	}
}