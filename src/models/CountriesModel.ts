import { WebPartContext } from "@microsoft/sp-webpart-base";
import { APIResponse } from "./ApiModel";

export interface CountriesFormData {

}

export interface SubmitCountriesFormData {
}

export interface ICountriesFormProps {
	description: string;
	context: WebPartContext;
}

export interface ICountriesFormState {
	form: CountriesFormData;
	apiMessage?: APIResponse;
	errors: { [key: string]: string };
	isSubmitting: boolean;
	PopUpWindowCloseButton: boolean;
}

export interface ICountriesViewProps {
	form: CountriesFormData;
	apiMessage?: APIResponse;
	errors: { [key: string]: string };
	onSubmit: (e: React.FormEvent<HTMLFormElement>) => void;
	onCancel: () => void;
	onInputChange: (name: string, value: any) => void;
	closePopUpWindow: () => void;
	PopUpWindowCloseButton: boolean;
	isSubmitting: boolean;
}