import BaseApi from "./BaseApi";
import { preMadeElectionsFormDataMock } from "../mock/elections/preMadeElectionsFormData";
import { APIResponse } from "../models/ApiModel";
import { SubmitCountriesFormData, CountriesFormData } from "../models/CountriesModel";
import { formatDateTimeForForm } from "../utils/dateUtils";
import { sp } from 'sp-pnp-js';
import { YesNoToBoolean } from "../utils/BooleanUtils";

export default class CountriesAPI extends BaseApi {
	public async submitElectionsForm(formData: SubmitCountriesFormData): Promise<APIResponse> {
		const payload = {
			__metadata: { type: "SP.Data.ContractdatabaseListItem" },
			...formData
		};

		return this.handleRequest(() =>
			this.candidaturePlatformApiClient.post("/_api/web/lists/getByTitle('Contract Database')/items", payload)
		);
	}

	public async editElectionForm(formData: SubmitCountriesFormData, id: number): Promise<APIResponse> {
		const payload = {
			__metadata: { type: "SP.Data.ContractdatabaseListItem" },
			...formData
		};

		return this.handleRequest(() =>
			this.candidaturePlatformApiClient.post(
				`/_api/web/lists/getByTitle('Elections Database')/items(${id})`,
				payload,
				{
					headers: {
						"X-HTTP-Method": "MERGE",
						"IF-MATCH": "*"
					}
				}
			)
		);
	}

	public async getElectionFormById(id?: number): Promise<CountriesFormData | null> {
		if (!id) return null;

		try {
			const item = await sp.web.lists
				.getByTitle("Elections")
				.items.getById(id)
				.select(
					"Election_x0020_body/Id",
					"Election_x0020_body/Title",
					"Post/Id",
					"Post/Title",
					"Date",
					"IsTentativeElectionDate",
					"CandidatureAnnouncementDeadline",
					"Seats",
					"IsTentativeSeatCount",
					"EligibleForVoteSwaps",
					"DanishVotes",
					"ResponsibleLineAuthority/Id",
					"ResponsibleLineAuthority/Title",
				)
				.expand(
					"Election_x0020_body",
					"Post",
					"ResponsibleLineAuthority",
				)
				.get();

			const candidaturesForm: CountriesFormData = {
				Body: item.Body,
				Post: item.Post,
				RegionalGroup: item.RegionalGroup,
				Status: item.Status,
				ElectionDate: formatDateTimeForForm(item.ElectionDate),
				TentativeDate: YesNoToBoolean(item.TentativeDate),
				AnnouncementDeadline: formatDateTimeForForm(item.AnnouncementDeadline),
				Seats: item.Seats,
				TentativeSeatCount: YesNoToBoolean(item.TentativeSeatCount),
				EligibleForVoteSwaps: YesNoToBoolean(item.EligibleForVoteSwaps),
				DanishVotesInElection: item.DanishVotesInElection,
				ResponsibleLineAuthorities: item.ResponsibleLineAuthorities
			};

			return candidaturesForm;
		} catch (error) {
			console.error("Error fetching contract item:", error);
			return null;
		}
	}

	public async getCountriesFormByIdMock(id?: number): Promise<CountriesFormData> {
		return preMadeElectionsFormDataMock;
	}
}