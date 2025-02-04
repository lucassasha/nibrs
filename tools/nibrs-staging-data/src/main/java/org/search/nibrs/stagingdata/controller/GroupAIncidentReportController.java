/*
 * Copyright 2016 SEARCH-The National Consortium for Justice Information and Statistics
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.search.nibrs.stagingdata.controller;

import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.search.nibrs.model.GroupAIncidentReport;
import org.search.nibrs.stagingdata.model.search.IncidentDeleteRequest;
import org.search.nibrs.stagingdata.model.segment.AdministrativeSegment;
import org.search.nibrs.stagingdata.service.GroupAIncidentService;
import org.search.nibrs.stagingdata.util.BaselineIncidentFactory;
import org.search.nibrs.util.CustomPair;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class GroupAIncidentReportController {
	@SuppressWarnings("unused")
	private static final Log log = LogFactory.getLog(GroupAIncidentReportController.class);

	@Autowired
	private GroupAIncidentService groupAIncidentService;
	
	@RequestMapping("/groupAIncidentReports")
	public List<AdministrativeSegment> getAllGroupAIncidentReport(){
		return groupAIncidentService.findAllAdministrativeSegments();
	}
	
	@RequestMapping("/groupAIncidentReport")
	public GroupAIncidentReport getGroupAIncidentReport(){
		return BaselineIncidentFactory.getBaselineIncident();
	}
	
	@RequestMapping(value="/groupAIncidentReports", method=RequestMethod.POST)
	public void save(@RequestBody List<GroupAIncidentReport> groupAIncidentReports){
		groupAIncidentService.saveGroupAIncidentReports(groupAIncidentReports.toArray(new GroupAIncidentReport[groupAIncidentReports.size()]));
	}
	
	@RequestMapping(value="/groupAIncidentReportsToXml", method=RequestMethod.POST)
	public void convert(@RequestBody CustomPair<String, List<GroupAIncidentReport>> groupAIncidentReportsPair) throws Exception{
		groupAIncidentService.convertAndWriteGroupAIncidentReports(groupAIncidentReportsPair);
	}
	
	@RequestMapping(value="/groupAIncidentReports/{incidentNumber}", method=RequestMethod.DELETE)
	public void deleteReport(@PathVariable("incidentNumber") String incidentNumber){
		groupAIncidentService.deleteGroupAIncidentReport(incidentNumber);
	}

	@RequestMapping(value="/groupAIncidentReports/{ori}/{yearOfTape}/{monthOfTape}", method=RequestMethod.DELETE)
	public @ResponseBody String deleteReportBySubmissionDate(@PathVariable("ori") String ori, @PathVariable("yearOfTape") String yearOfTape, @PathVariable("monthOfTape") String monthOfTape){
		Integer count = groupAIncidentService.deleteGroupAIncidentReports(ori, yearOfTape, monthOfTape);
		
		return String.valueOf(count) + " records are deleted";
	}
	
	@RequestMapping(value="/groupAIncidentReports/administrativeSegmentIds", method=RequestMethod.GET)
	public @ResponseBody List<Integer> getAdministrativeSegmentIdsByIncidentDeleteRequest(@RequestBody IncidentDeleteRequest incidentDeleteRequest){
		List<Integer> administrativeSegmentIds = groupAIncidentService.findAdministrativeSegmentIdsByIncidentDeleteRequest(incidentDeleteRequest);
		
		return administrativeSegmentIds;
	}
	
	@RequestMapping(value="/groupAIncidentReports", method=RequestMethod.DELETE)
	public @ResponseBody String deleteByIncidentDeleteRequest(@RequestBody IncidentDeleteRequest incidentDeleteRequest){
		Integer count = groupAIncidentService.deleteGroupAIncidentReportsByRequest(incidentDeleteRequest);
		
		return String.valueOf(count) + " group A incident reports are deleted.";
	}
	
}
