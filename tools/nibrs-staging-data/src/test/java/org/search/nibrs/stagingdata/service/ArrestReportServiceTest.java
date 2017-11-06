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
package org.search.nibrs.stagingdata.service;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.junit.Assert.assertThat;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.HashSet;
import java.util.Set;

import org.apache.commons.lang3.time.DateUtils;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.search.nibrs.stagingdata.model.Agency;
import org.search.nibrs.stagingdata.model.ArrestReportSegmentWasArmedWith;
import org.search.nibrs.stagingdata.model.DateType;
import org.search.nibrs.stagingdata.model.DispositionOfArresteeUnder18Type;
import org.search.nibrs.stagingdata.model.EthnicityOfPersonType;
import org.search.nibrs.stagingdata.model.RaceOfPersonType;
import org.search.nibrs.stagingdata.model.ResidentStatusOfPersonType;
import org.search.nibrs.stagingdata.model.SegmentActionTypeType;
import org.search.nibrs.stagingdata.model.SexOfPersonType;
import org.search.nibrs.stagingdata.model.TypeOfArrestType;
import org.search.nibrs.stagingdata.model.UcrOffenseCodeType;
import org.search.nibrs.stagingdata.model.segment.ArrestReportSegment;
import org.search.nibrs.stagingdata.repository.AgencyRepository;
import org.search.nibrs.stagingdata.repository.ArresteeWasArmedWithTypeRepository;
import org.search.nibrs.stagingdata.repository.DateTypeRepository;
import org.search.nibrs.stagingdata.repository.DispositionOfArresteeUnder18TypeRepository;
import org.search.nibrs.stagingdata.repository.EthnicityOfPersonTypeRepository;
import org.search.nibrs.stagingdata.repository.RaceOfPersonTypeRepository;
import org.search.nibrs.stagingdata.repository.ResidentStatusOfPersonTypeRepository;
import org.search.nibrs.stagingdata.repository.SegmentActionTypeRepository;
import org.search.nibrs.stagingdata.repository.SexOfPersonTypeRepository;
import org.search.nibrs.stagingdata.repository.TypeOfArrestTypeRepository;
import org.search.nibrs.stagingdata.repository.UcrOffenseCodeTypeRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ArrestReportServiceTest {

	@Autowired
	public ArrestReportService arrestReportService; 
	@Autowired
	public DateTypeRepository dateTypeRepository; 
	@Autowired
	public UcrOffenseCodeTypeRepository ucrOffenseCodeTypeRepository; 
	@Autowired
	public AgencyRepository agencyRepository; 
	@Autowired
	public DispositionOfArresteeUnder18TypeRepository dispositionOfArresteeUnder18TypeRepository; 
	@Autowired
	public EthnicityOfPersonTypeRepository ethnicityOfPersonTypeRepository; 
	@Autowired
	public RaceOfPersonTypeRepository raceOfPersonTypeRepository; 
	@Autowired
	public SexOfPersonTypeRepository sexOfPersonTypeRepository; 
	@Autowired
	public SegmentActionTypeRepository segmentActionTypeRepository; 
	@Autowired
	public TypeOfArrestTypeRepository typeOfArrestTypeRepository; 
	@Autowired
	public ResidentStatusOfPersonTypeRepository residentStatusOfPersonTypeRepository; 
	@Autowired
	public ArresteeWasArmedWithTypeRepository arresteeWasArmedWithTypeRepository; 
	
	@Test
	public void test() {
		ArrestReportSegment arrestReportSegment = getBasicArrestReportSegment();
		arrestReportService.saveArrestReportSegment(arrestReportSegment); 
		
		ArrestReportSegment persisted = 
				arrestReportService.findArrestReportSegment(arrestReportSegment.getArrestReportSegmentId());
		assertThat(persisted.getAgeOfArresteeMax(), equalTo(25));
		assertThat(persisted.getAgeOfArresteeMin(), equalTo(22));
		assertTrue(DateUtils.isSameDay(persisted.getArrestDate(), Date.from(LocalDateTime.of(2016, 6, 12, 10, 7, 46).atZone(ZoneId.systemDefault()).toInstant())));
		
		assertThat(persisted.getArrestDateType().getDateTypeId(), equalTo(2355));
		assertThat(persisted.getArrestDateType().getDateMMDDYYYY(), equalTo("06122016"));
		
		assertThat(persisted.getArresteeSequenceNumber(), equalTo(1));
		assertThat(persisted.getAgency().getAgencyId(), equalTo(1));
		assertThat(persisted.getArrestTransactionNumber(), equalTo("arrestTr"));
		assertThat(persisted.getCityIndicator(), equalTo("Y"));
		assertThat(persisted.getDispositionOfArresteeUnder18Type().getDispositionOfArresteeUnder18TypeId(), equalTo(2));
		assertThat(persisted.getEthnicityOfPersonType().getEthnicityOfPersonTypeId(), equalTo(1));
		assertThat(persisted.getMonthOfTape(), equalTo("12"));
		assertThat(persisted.getOri(), equalTo("ori"));;
		assertThat(persisted.getRaceOfPersonType().getRaceOfPersonCode(), equalTo("W"));
		assertThat(persisted.getResidentStatusOfPersonType().getResidentStatusOfPersonTypeId(), equalTo(1));
		assertThat(persisted.getSegmentActionType().getSegmentActionTypeTypeId(), equalTo(1));
		assertThat(persisted.getSexOfPersonType().getSexOfPersonCode(), equalTo("F"));
		assertThat(persisted.getTypeOfArrestType().getTypeOfArrestTypeId(), equalTo(1));
		assertThat(persisted.getUcrOffenseCodeType().getUcrOffenseCodeTypeId(), equalTo(520));
		assertThat(persisted.getYearOfTape(), equalTo("2016"));
		
		Set<ArrestReportSegmentWasArmedWith> arrestReportSegmentWasArmedWiths = 
				persisted.getArrestReportSegmentWasArmedWiths();
		assertThat(arrestReportSegmentWasArmedWiths.size(), equalTo(2));
		
		for (ArrestReportSegmentWasArmedWith reportSegmentWasArmedWith: arrestReportSegmentWasArmedWiths){
			if (reportSegmentWasArmedWith.getAutomaticWeaponIndicator().equals("A")){
				assertThat(reportSegmentWasArmedWith.getArresteeWasArmedWithType().getArresteeWasArmedWithCode(), equalTo("12"));
			}
			else if (reportSegmentWasArmedWith.getAutomaticWeaponIndicator().equals("")){
				assertThat(reportSegmentWasArmedWith.getArresteeWasArmedWithType().getArresteeWasArmedWithCode(), equalTo("11"));
			}
			else {
				fail("Unexpected reportSegmentWasArmedWith.getAutomaticWeaponIndicator() value");
			}
		}

	}
	
	@SuppressWarnings("serial")
	public ArrestReportSegment getBasicArrestReportSegment(){
		ArrestReportSegment arrestReportSegment = new ArrestReportSegment();
		arrestReportSegment.setAgeOfArresteeMax(25);
		arrestReportSegment.setAgeOfArresteeMin(22);
		arrestReportSegment.setArrestDate(Date.from(LocalDateTime.of(2016, 6, 12, 10, 7, 46).atZone(ZoneId.systemDefault()).toInstant()));
		
		DateType arrestDateType = dateTypeRepository.findFirstByDateMMDDYYYY("06122016");
		arrestReportSegment.setArrestDateType(arrestDateType);
		
		arrestReportSegment.setArresteeSequenceNumber(1);
		
		Agency agency = agencyRepository.findFirstByAgencyOri("agencyORI");
		arrestReportSegment.setAgency(agency);
		arrestReportSegment.setArrestTransactionNumber("arrestTr");
		arrestReportSegment.setCityIndicator("Y");
		
		DispositionOfArresteeUnder18Type dispositionOfArresteeUnder18Type
			= dispositionOfArresteeUnder18TypeRepository.findFirstByDispositionOfArresteeUnder18Code("H");
		arrestReportSegment.setDispositionOfArresteeUnder18Type(dispositionOfArresteeUnder18Type);
		
		EthnicityOfPersonType ethnicityOfPersonType = ethnicityOfPersonTypeRepository.findFirstByEthnicityOfPersonCode("N");
		arrestReportSegment.setEthnicityOfPersonType(ethnicityOfPersonType);
		arrestReportSegment.setMonthOfTape("12");
		arrestReportSegment.setOri("ori");;
		
		RaceOfPersonType raceOfPersonType = raceOfPersonTypeRepository.findFirstByRaceOfPersonCode("W");
		arrestReportSegment.setRaceOfPersonType(raceOfPersonType);
		
		ResidentStatusOfPersonType residentStatusOfPersonType = residentStatusOfPersonTypeRepository.findFirstByResidentStatusOfPersonCode("N");
		arrestReportSegment.setResidentStatusOfPersonType(residentStatusOfPersonType);
		
		SegmentActionTypeType segmentActionTypeType = segmentActionTypeRepository.findFirstBySegmentActionTypeCode("I");
		arrestReportSegment.setSegmentActionType(segmentActionTypeType);
		
		SexOfPersonType sexOfPersonType = sexOfPersonTypeRepository.findFirstBySexOfPersonCode("F");
		arrestReportSegment.setSexOfPersonType(sexOfPersonType);
		
		TypeOfArrestType typeOfArrestType = typeOfArrestTypeRepository.findFirstByTypeOfArrestCode("O");
		arrestReportSegment.setTypeOfArrestType(typeOfArrestType);
		UcrOffenseCodeType ucrOffenseCode = ucrOffenseCodeTypeRepository.findFirstByUcrOffenseCode("520");
		arrestReportSegment.setUcrOffenseCodeType(ucrOffenseCode);
		arrestReportSegment.setYearOfTape("2016");
		
		ArrestReportSegmentWasArmedWith arrestReportSegmentWasArmedWith1 = new ArrestReportSegmentWasArmedWith();
		arrestReportSegmentWasArmedWith1.setArrestReportSegment(arrestReportSegment);
		arrestReportSegmentWasArmedWith1.setAutomaticWeaponIndicator("A");
		arrestReportSegmentWasArmedWith1.setArresteeWasArmedWithType(arresteeWasArmedWithTypeRepository.findOne(120));
		
		ArrestReportSegmentWasArmedWith arrestReportSegmentWasArmedWith2 = new ArrestReportSegmentWasArmedWith();
		arrestReportSegmentWasArmedWith2.setArrestReportSegment(arrestReportSegment);
		arrestReportSegmentWasArmedWith2.setAutomaticWeaponIndicator("");
		arrestReportSegmentWasArmedWith2.setArresteeWasArmedWithType(arresteeWasArmedWithTypeRepository.findOne(110));
		
		arrestReportSegment.setArrestReportSegmentWasArmedWiths(new HashSet<ArrestReportSegmentWasArmedWith>(){{
			add(arrestReportSegmentWasArmedWith1);
			add(arrestReportSegmentWasArmedWith2);
		}});
		
		return arrestReportSegment; 
	}

}