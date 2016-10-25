package org.search.nibrs.validation.groupa;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.search.nibrs.common.NIBRSError;
import org.search.nibrs.model.NIBRSAge;
import org.search.nibrs.model.VictimSegment;
import org.search.nibrs.model.codes.AdditionalJustifiableHomicideCircumstancesCode;
import org.search.nibrs.model.codes.AggravatedAssaultHomicideCircumstancesCode;
import org.search.nibrs.model.codes.EthnicityOfVictim;
import org.search.nibrs.model.codes.NIBRSErrorCode;
import org.search.nibrs.model.codes.OffenseCode;
import org.search.nibrs.model.codes.RaceOfVictimCode;
import org.search.nibrs.model.codes.RelationshipOfVictimToOffenderCode;
import org.search.nibrs.model.codes.ResidentStatusCode;
import org.search.nibrs.model.codes.SexOfVictimCode;
import org.search.nibrs.model.codes.TypeOfOfficerActivityCircumstance;
import org.search.nibrs.model.codes.TypeOfVictimCode;
import org.search.nibrs.validation.rules.DuplicateCodedValueRule;
import org.search.nibrs.validation.rules.NotAllBlankRule;
import org.search.nibrs.validation.rules.NotBlankRule;
import org.search.nibrs.validation.rules.Rule;
import org.search.nibrs.validation.rules.ValidValueListRule;


public class VictimSegmentRulesFactory {
		

	private List<Rule<VictimSegment>> rulesList;

	
	public static VictimSegmentRulesFactory instance(){
		
		return new VictimSegmentRulesFactory();
	}
	
	private static final List<String> PERSON_OFFENSE_LIST = Arrays.asList(OffenseCode._13A.code, OffenseCode._13B.code, OffenseCode._13C.code, 
			OffenseCode._09A.code, OffenseCode._09B.code, OffenseCode._64A.code, OffenseCode._64B.code, OffenseCode._100.code,
			OffenseCode._11A.code, OffenseCode._11B.code, OffenseCode._11C.code, OffenseCode._11D.code, OffenseCode._36A.code,
			OffenseCode._36B.code);
	
	private static final List<String> PROPERTY_CRIM_LIST = Arrays.asList(OffenseCode._200.code,
			OffenseCode._510.code, OffenseCode._220.code, OffenseCode._250.code,
			OffenseCode._290.code, OffenseCode._270.code, OffenseCode._210.code,
			OffenseCode._26A.code, OffenseCode._26B.code, OffenseCode._26C.code,
			OffenseCode._26D.code, OffenseCode._26E.code, OffenseCode._26F.code,
			OffenseCode._26G.code, OffenseCode._23A.code, OffenseCode._23B.code,
			OffenseCode._23C.code, OffenseCode._23D.code, OffenseCode._23E.code,
			OffenseCode._23F.code, OffenseCode._23G.code, OffenseCode._23H.code,
			OffenseCode._240.code, OffenseCode._120.code, OffenseCode._280.code					
	);		
	
	private VictimSegmentRulesFactory() {
		
		rulesList = new ArrayList<Rule<VictimSegment>>();
		
		initRules(rulesList);
	}
	
	private void initRules(List<Rule<VictimSegment>> rulesList){
		
		rulesList.add(getRule401ForSequenceNumber());
		
		rulesList.add(getRule401ForVictimConnectedToUcrOffenseCode());
		
		rulesList.add(getRule406ForVictimConnectedToUcrOffenseCode());
				
		rulesList.add(getRule401ForTypeOfVictim());
		
		rulesList.add(getRule404ForTypeOfOfficerActivityCircumstance());
		
		rulesList.add(getRule404ForOfficerAssignmentType());
				
		rulesList.add(getRule404ForOfficerOriOtherJurisdiction());
		
		rulesList.add(getRule404ForAgeOfVictim());
		
		rulesList.add(getRule410ForAgeOfVictim());
		
		rulesList.add(getRule422ForAgeOfVictim());
		
		rulesList.add(getRule458ForAgeOfVictim());
		
		rulesList.add(getRule450ForAgeOfVictim());
		
		rulesList.add(getRule453ForAgeOfVictim());
		
		rulesList.add(getRule481ForAgeOfVictim());
		
		rulesList.add(getRule404ForSexOfVictim());	
		
		rulesList.add(getRule469ForSexOfVictim());
		
		rulesList.add(getRule458ForSexOfVictim());		
		
		rulesList.add(getRule404ForRaceOfVictim());		
		
		rulesList.add(getRule404ForEthnicityOfVictim());
		
		rulesList.add(getRule458ForEthnicityOfVictim());
		
		rulesList.add(getRule458ForRaceOfVictim());
		
		rulesList.add(getRule404ForResidentStatusOfVictim());
		
		rulesList.add(getRule458ForResidentStatusOfVictim());
		
		rulesList.add(getRule404ForAggravatedAssaultHomicideCircumstances());
		
		rulesList.add(getRule406ForAggravatedAssaultHomicideCircumstances());
		
		rulesList.add(getRule404ForAdditionalJustifiablehomicideCircsumstances());
		
		rulesList.add(getRule404ForTypeOfInjury());
		
		rulesList.add(getRule406ForTypeOfInjury());
		
		rulesList.add(getRule404OffenderNumberToBeRelated());
		
		rulesList.add(getRule404OffenderNumberToBeRelated());
		
		rulesList.add(getRule458ForOffenderNumberToBeRelated());
		
		rulesList.add(getRule404ForRelationshipOfVictimToOffender());
		
		rulesList.add(geRule465ForTypeOfVictim());
		
		rulesList.add(getRule467ForTypeOfVictim());
		
		rulesList.add(getRule482ForTypeOfVictim());
		
		rulesList.add(getRule454ForTypeOfOfficerActivityCircumstance());
		
		rulesList.add(getRule453ForSexOfVictim());
		
		rulesList.add(getRule464ForTypeOfVictim());
		
		rulesList.add(getRule483ForTypeOfOfficerActivity());
		
		rulesList.add(getRule453ForRaceOfVictim());
		
		rulesList.add(getRule468ForRelationshipOfVictimToOffender());
	}
	
		
	public List<Rule<VictimSegment>> getRulesList() {
		return rulesList;
	}
		
		
	/**
	 * (Victim Sequence Number The referenced data element in a Group A Incident
	 * Report must be populated with a valid data value and cannot be blank.
	 */
	public Rule<VictimSegment> getRule401ForSequenceNumber(){
		
		Rule<VictimSegment> sequenceNumberNotEmptyRule = new Rule<VictimSegment>(){

			@Override
			public NIBRSError apply(VictimSegment subject) {

				Integer iSeqNum = subject.getVictimSequenceNumber();
				
				NIBRSError rSeqNumInvalidError = null;
				
				if(iSeqNum ==  null || iSeqNum < 1 || iSeqNum > 999){
					
					rSeqNumInvalidError = subject.getErrorTemplate();				
					rSeqNumInvalidError.setNIBRSErrorCode(NIBRSErrorCode._401);					
					rSeqNumInvalidError.setDataElementIdentifier("23");															
				}				
				return rSeqNumInvalidError;
			}};		
									
		return sequenceNumberNotEmptyRule;
	}
	
					
	/**
	 * (Victim Connected to UCR Offense Code) The referenced data element in a
	 * Group A Incident Report must be populated with a valid data value and
	 * cannot be blank.
	 */
	public Rule<VictimSegment> getRule401ForVictimConnectedToUcrOffenseCode() {		
		
		// TODO confirm if value should be checked against OffenseCode enum
		
		NotAllBlankRule<VictimSegment> ucrOffenseNotAllBlankRule = new NotAllBlankRule<VictimSegment>("ucrOffenseCodeConnection", 
				"24", VictimSegment.class, NIBRSErrorCode._401);		
		
		return ucrOffenseNotAllBlankRule;
	}	
	
	
	/**
	 * (Type of Victim) The referenced data element in a Group A Incident Report
	 * must be populated with a valid data value and cannot be blank.
	 */
	public Rule<VictimSegment> getRule401ForTypeOfVictim(){
								
		ValidValueListRule<VictimSegment> typeOfVictimValidValue401Rule = new ValidValueListRule<VictimSegment>(
				"typeOfVictim", "25", VictimSegment.class, NIBRSErrorCode._401, 
				TypeOfVictimCode.codeSet());
		
		return typeOfVictimValidValue401Rule;
	}

	/**
	 * (Type of Officer Activity/Circumstance) The referenced data element in a
	 * Group A Incident Report must be populated with a valid data value and
	 * cannot be blank.
	 */
	public Rule<VictimSegment> getRule404ForTypeOfOfficerActivityCircumstance(){
		
		ValidValueListRule<VictimSegment> validValueListRule = new ValidValueListRule<VictimSegment>(
				"typeOfOfficerActivityCircumstance", "25A", VictimSegment.class, NIBRSErrorCode._404, 
				TypeOfOfficerActivityCircumstance.codeSet());
		
		return validValueListRule;
	}


	/**
	 * (Sex of Victim) The referenced data element in a Group A Incident Report
	 * must be populated with a valid data value and cannot be blank.
	 */
	public Rule<VictimSegment> getRule404ForSexOfVictim(){
		
		ValidValueListRule<VictimSegment> validValueListRule = new ValidValueListRule<VictimSegment>(
				"sex", "27", VictimSegment.class, NIBRSErrorCode._404, SexOfVictimCode.codeSet());
		
		return validValueListRule;
	}

	/**
	 * (Officer–ORI Other Jurisdiction) The referenced data element in a Group A
	 * Incident Report must be populated with a valid data value and cannot be
	 * blank.
	 */
	public Rule<VictimSegment> getRule404ForOfficerOriOtherJurisdiction(){
						
		Rule<VictimSegment> officerOriRule = new Rule<VictimSegment>(){

			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
								
				String officerOri = victimSegment.getOfficerOtherJurisdictionORI();
				
				boolean oriValid = true;
				
				if(StringUtils.isNotEmpty(officerOri)){

					if(StringUtils.isNumeric(officerOri)){
						
						int iOri = Integer.parseInt(officerOri);
						
						if(iOri < 0 || iOri > 9){
							
							oriValid = false;
						}						
					}else{
						
						oriValid = StringUtils.isAlpha(officerOri) && StringUtils.isAllUpperCase(officerOri);						
					}					
				}
				
				if(!oriValid){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setDataElementIdentifier("25C");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._404);
				}								
				return rNIBRSError;
			}			
		};		
		return officerOriRule;						
	}

	/**
	 * Officer Assignment Type) The referenced data element in a Group A
	 * Incident Report must be populated with a valid data value and cannot be
	 * blank.
	 */
	public Rule<VictimSegment> getRule404ForOfficerAssignmentType(){
		
		ValidValueListRule<VictimSegment> validValueListRule = new ValidValueListRule<VictimSegment>(
				"officerAssignmentType", "25B", VictimSegment.class, NIBRSErrorCode._404, 
				TypeOfOfficerActivityCircumstance.codeSet());
		
		return validValueListRule;
	}

	
	/**
	 * (Ethnicity of Victim) The referenced data element in a Group A Incident
	 * Report must be populated with a valid data value and cannot be blank
	 */
	public Rule<VictimSegment> getRule404ForEthnicityOfVictim(){
		
		//TODO account for element being optional based on typeOfVicitmValue  
		
		ValidValueListRule<VictimSegment> validValueListRule = new ValidValueListRule<VictimSegment>(
				"ethnicity", "29", VictimSegment.class, NIBRSErrorCode._404, EthnicityOfVictim.codeSet());
		
		return validValueListRule;
	}

	
	public Rule<VictimSegment> getRule404ForTypeOfInjury(){
		
		NotAllBlankRule<VictimSegment> notAllBlankRule = new NotAllBlankRule<VictimSegment>("typeOfInjury", 
				"typeOfInjury", VictimSegment.class, NIBRSErrorCode._404);
		
		return notAllBlankRule;
	}

	
	public Rule<VictimSegment> getRule404ForResidentStatusOfVictim(){
		
		ValidValueListRule<VictimSegment> validValueListRule = new ValidValueListRule<VictimSegment>(
				"residentStatusOfVictim", "30", VictimSegment.class, NIBRSErrorCode._404, 
				ResidentStatusCode.codeSet());
		
		return validValueListRule;
	}

	
	public Rule<VictimSegment> getRule404OffenderNumberToBeRelated(){
		
		//TODO add condition		
		NotAllBlankRule<VictimSegment> notAllBlankRule = new NotAllBlankRule<VictimSegment>("offenderNumberRelated", 
				"34", VictimSegment.class, NIBRSErrorCode._404);
		
		return notAllBlankRule;
	}

	/**
	 * (Aggravated Assault/Homicide Circumstances) The referenced data element
	 * in a Group A Incident Report must be populated with a valid data value
	 * and cannot be blank.
	 */
	public Rule<VictimSegment> getRule404ForAggravatedAssaultHomicideCircumstances(){
		
		ValidValueListRule<VictimSegment> validValueListRule = new ValidValueListRule<VictimSegment>("aggravatedAssaultHomicideCircumstances", 
				"31", VictimSegment.class, NIBRSErrorCode._404, AggravatedAssaultHomicideCircumstancesCode.codeSet());
		
		return validValueListRule;
	}

	public Rule<VictimSegment> getRule404ForRelationshipOfVictimToOffender(){
			
		ValidValueListRule<VictimSegment> validValueListRule = new ValidValueListRule<VictimSegment>("victimOffenderRelationship", 
				"35", VictimSegment.class, NIBRSErrorCode._404, RelationshipOfVictimToOffenderCode.codeSet());
		
		return validValueListRule;		
	}

	public Rule<VictimSegment> getRule404ForAdditionalJustifiablehomicideCircsumstances(){
		
		ValidValueListRule<VictimSegment> validValueListRule = new ValidValueListRule<VictimSegment>(
				"additionalJustifiableHomicideCircumstances", "32", VictimSegment.class, 
				NIBRSErrorCode._404, AdditionalJustifiableHomicideCircumstancesCode.codeSet());
		
		return validValueListRule;
	}

	public Rule<VictimSegment> getRule404ForRaceOfVictim(){
	
		ValidValueListRule<VictimSegment> validValueListRule = new ValidValueListRule<VictimSegment>(
				"race", "28", VictimSegment.class, NIBRSErrorCode._404, RaceOfVictimCode.codeSet());
		
		return validValueListRule;
	}

	public Rule<VictimSegment> getRule404ForAgeOfVictim(){
		
		//TODO maybe use AgeOfVictimCode enum to validate values. range is tricky
		
		NotBlankRule<VictimSegment> notBlankRule = new NotBlankRule<VictimSegment>(
				"age", "26", VictimSegment.class, NIBRSErrorCode._404);
		
		return notBlankRule;
	}

	// (Victim Connected to UCR Offense Code) The referenced data element in
	// error is one that contains multiple data values. When more than one code
	// is entered, none can be duplicate codes.
	public Rule<VictimSegment> getRule406ForVictimConnectedToUcrOffenseCode(){
		
		DuplicateCodedValueRule<VictimSegment> duplicateCodedValueRule = new DuplicateCodedValueRule<>
			("ucrOffenseCodeConnection", "24", VictimSegment.class, NIBRSErrorCode._406);
		
		return duplicateCodedValueRule;
	}	
		
	
	/**
	 * (Aggravated Assault/Homicide Circumstances The referenced data element in
	 * error is one that contains multiple data values. When more than one code
	 * is entered, none can be duplicate codes.
	 */
	public Rule<VictimSegment> getRule406ForAggravatedAssaultHomicideCircumstances(){
	
		DuplicateCodedValueRule<VictimSegment> duplicateValueRule = new DuplicateCodedValueRule<VictimSegment>(
				"aggravatedAssaultHomicideCircumstances", "31", VictimSegment.class, NIBRSErrorCode._406);
		
		return duplicateValueRule;		
	}

	public Rule<VictimSegment> getRule406OffenderNumberToBeRelated(){
		
		DuplicateCodedValueRule<VictimSegment> noDuplicateRule = new DuplicateCodedValueRule<VictimSegment>("offenderNumberRelated", 
				"34", VictimSegment.class, NIBRSErrorCode._406);
		
		return noDuplicateRule;
	}

	public Rule<VictimSegment> getRule406ForTypeOfInjury(){
		
		DuplicateCodedValueRule<VictimSegment> duplicateCodedValueRule = new DuplicateCodedValueRule<VictimSegment>(
				"typeOfInjury",	"33", VictimSegment.class, NIBRSErrorCode._406);
		
		return duplicateCodedValueRule;
	}

	public Rule<VictimSegment> getRule410ForAgeOfVictim(){
		
		Rule<VictimSegment> ageRangeOrderRule = new Rule<VictimSegment>(){
			
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
	
				NIBRSAge nibrsAge = victimSegment.getAge();
				
				Integer ageMin = nibrsAge.getAgeMin();
				Integer ageMax = nibrsAge.getAgeMax();
				
				if(ageMin != null && ageMax != null && ageMin > ageMax){
					
					rNIBRSError = victimSegment.getErrorTemplate();					
					rNIBRSError.setDataElementIdentifier("26");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._410);					
				}								
				return rNIBRSError;
			}
		};
		return ageRangeOrderRule;
	}

	
	/**
	 * Data Element 33 (Type Injury) can only be entered when one or more of the
	 * offenses in Data Element 24 (Victim Connected to UCR Offense Code) are:
	 * 100=Kidnapping/Abduction 11A=Rape 11B=Sodomy 11C=Sexual Assault With An
	 * Object 11D=Fondling 120=Robbery 13A=Aggravated Assault 13B=Simple Assault
	 * 210=Extortion/Blackmail 64A=Human Trafficking, Commercial Sex Acts
	 * 64B=Human Trafficking, Involuntary Servitude
	 */
	public Rule<VictimSegment> getRule419ForTypeOfInjury(){
		
		Rule<VictimSegment> typeInjuryForNonPersonRule = new Rule<VictimSegment>(){
	
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
								
				List<String> offenseList =  victimSegment.getUcrOffenseCodeList();
				
				if(offenseList != null && !offenseList.isEmpty() 
					&& !CollectionUtils.containsAny(offenseList, PERSON_OFFENSE_LIST)){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setDataElementIdentifier("33");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._419);
				}
				
				return rNIBRSError;
			}						
		};		
		return typeInjuryForNonPersonRule;
	}

	
	/**
	 * (Aggravated Assault/Homicide Circumstances) Circumstances) can only be
	 * entered when one or more of the offenses in Data Element 24 (Victim
	 * Connected to UCR Offense Code) are: 09A=Murder and Non-negligent
	 * Manslaughter 09B=Negligent Manslaughter 09C=Justifiable Homicide
	 * 13A=Aggravated Assault
	 */
	public Rule<VictimSegment> getRule419ForAggravatedAssaultHomicideCircumstances(){
		
		Rule<VictimSegment> assaultRule = new Rule<VictimSegment>(){
	
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
	
				NIBRSError rNIBRSError = null;
								
				List<String> assaultCircList = victimSegment.getAggravatedAssaultHomicideCircumstancesList();
				
				boolean hasAssaultEntered = assaultCircList != null && !assaultCircList.isEmpty();
				
				
				List<String> ucrOffenseList = victimSegment.getUcrOffenseCodeList();
				
				List<String> offenseListForAssault = Arrays.asList(OffenseCode._09A.code,
						OffenseCode._09B.code, OffenseCode._09C.code, OffenseCode._13A.code);
				
				boolean hasAssaultOffense = CollectionUtils.containsAny(ucrOffenseList, offenseListForAssault);
				
				if(hasAssaultEntered && !hasAssaultOffense){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._419);
					rNIBRSError.setDataElementIdentifier("31");
				}
				
				return rNIBRSError;
			}						
		};
		return assaultRule;
	}

	
	public Rule<VictimSegment> getRule422ForAgeOfVictim(){
		
		Rule<VictimSegment> ageMin00Rule = new Rule<VictimSegment>(){
			
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
	
				NIBRSAge nibrsAge = victimSegment.getAge();
								
				if(nibrsAge.isAgeRange() && nibrsAge.getAgeMin() != null 
						&& nibrsAge.getAgeMin() == 00){
	
					rNIBRSError = victimSegment.getErrorTemplate();					
					rNIBRSError.setDataElementIdentifier("26");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._422);
				}																	
				return rNIBRSError;
			}
		};
		return ageMin00Rule;
	}

	
	public Rule<VictimSegment> getRule450ForAgeOfVictim(){
		
		Rule<VictimSegment> ageOfSpouse = new Rule<VictimSegment>(){
			
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
				
				List<String> relationshipList = victimSegment.getVictimOffenderRelationshipList();
				
				if(relationshipList != null 
						&& relationshipList.contains(RelationshipOfVictimToOffenderCode.SP.code)){
				
					NIBRSAge nibrsAge = victimSegment.getAge();
					
					Integer minAge = nibrsAge.getAgeMin();
					
					if(minAge != null && minAge < 10){
	
						rNIBRSError = victimSegment.getErrorTemplate();					
						rNIBRSError.setDataElementIdentifier("26");
						rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._450);
					}									
				}
				return rNIBRSError;
			}
		};
		return ageOfSpouse;
	}


	
	/**
	 * (Race of Victim) The Data Element associated with this error must be
	 * present when Data Element 25 (Type of Victim) is I=Individual.
	 */
	public Rule<VictimSegment> getRule453ForRaceOfVictim(){
		
		Rule<VictimSegment> raceForIndividualRule = new Rule<VictimSegment>(){
	
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
	
				NIBRSError rNIBRSError = null;
				
				String race = victimSegment.getRace();
				
				boolean hasRace = StringUtils.isNotEmpty(race);
				
				String typeOfVictim = victimSegment.getTypeOfVictim();
				
				boolean isIndivideual = TypeOfVictimCode.I.equals(typeOfVictim);
				
				if(isIndivideual && !hasRace){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._453);
					rNIBRSError.setDataElementIdentifier("28");
				}				
				return rNIBRSError;
			}			
		};
		return raceForIndividualRule;
	}

	
	public Rule<VictimSegment> getRule453ForAgeOfVictim(){
		
		Rule<VictimSegment> ageOfIndividualRule = new Rule<VictimSegment>(){
	
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNibrsError = null;
								
				String victimType = victimSegment.getTypeOfVictim();
				
				boolean isIndividual = TypeOfVictimCode.I.code.equals(victimType);
				
				NIBRSAge victimAge = victimSegment.getAge();
				
				if(isIndividual && victimAge == null){
					
					rNibrsError = victimSegment.getErrorTemplate();
					
					rNibrsError.setNIBRSErrorCode(NIBRSErrorCode._453);
					rNibrsError.setDataElementIdentifier("26");
				}								
				return rNibrsError;
			}						
		}; 		
		return ageOfIndividualRule;		
	}

	
	public Rule<VictimSegment> getRule453ForSexOfVictim(){
		
		Rule<VictimSegment> sexOfIndividualRule = new Rule<VictimSegment>(){
	
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNibrsError = null;
								
				String victimType = victimSegment.getTypeOfVictim();
				
				boolean isIndividual = TypeOfVictimCode.I.code.equals(victimType);
				
				boolean containsSexCode = StringUtils.isNotEmpty(victimSegment.getSex());
				
				if(isIndividual && !containsSexCode){
					
					rNibrsError = victimSegment.getErrorTemplate();
					
					rNibrsError.setNIBRSErrorCode(NIBRSErrorCode._453);
					rNibrsError.setDataElementIdentifier("27");
				}								
				return rNibrsError;
			}						
		}; 		
		return sexOfIndividualRule;		
	}

	/**
	 * (Type of Officer Activity/Circumstance), Data Element 25B (Officer
	 * Assignment Type), Data Element 26 (Age of Victim), Data Element 27 (Sex
	 * of Victim), and Data Element 28 (Race of Victim) must be entered when
	 * Data Element 25 (Type of Victim) is L=Law Enforcement Officer.
	 */
	public Rule<VictimSegment> getRule454ForTypeOfOfficerActivityCircumstance(){
		
		Rule<VictimSegment> typeOfficerReqFieldsRule = new Rule<VictimSegment>(){
	
			NIBRSError rNibrsError = null;
			
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
	
				String victimType = victimSegment.getTypeOfVictim();
				
				boolean victimIsLawOfficer = TypeOfVictimCode.L.code.equals(victimType);
				
				if(victimIsLawOfficer){
										
					String circumstance = victimSegment.getTypeOfOfficerActivityCircumstance();
					
					String assignmentType = victimSegment.getOfficerAssignmentType();
					
					NIBRSAge age = victimSegment.getAge();
					
					String victimSex = victimSegment.getSex();
					
					String victimRace = victimSegment.getRace();	
					
					if(StringUtils.isEmpty(circumstance) 
							|| StringUtils.isEmpty(assignmentType)
							|| age == null
							|| StringUtils.isEmpty(victimSex)
							|| StringUtils.isEmpty(victimRace)){
						
						rNibrsError = victimSegment.getErrorTemplate();
						
						rNibrsError.setNIBRSErrorCode(NIBRSErrorCode._454);
						rNibrsError.setDataElementIdentifier("25A");						
					}					
				}								
				return rNibrsError;
			}
		};			
		return typeOfficerReqFieldsRule;
	}

	
	/**
	 * The Data Element associated with this error cannot be entered when Data
	 * Element 25 (Type of Victim) is not I=Individual or L=Law Enforcement
	 * Officer when Data Element 24 (Victim Connected to UCR Offense Code)
	 * contains a Crime Against Person.
	 */
	public Rule<VictimSegment> getRule458ForAgeOfVictim(){
		
		Rule<VictimSegment> ageForNonPersonRule = new Rule<VictimSegment>(){
									
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
			
				List<String> offenseList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasCrimeAgainstPerson = CollectionUtils.containsAny(offenseList, PERSON_OFFENSE_LIST);
								
				String sVictimType = victimSegment.getTypeOfVictim();
				
				boolean victimIsPerson = TypeOfVictimCode.I.code.equals(sVictimType)
						|| TypeOfVictimCode.L.code.equals(sVictimType);				
				
				NIBRSAge age = victimSegment.getAge();
				
				if(hasCrimeAgainstPerson && !victimIsPerson && age != null){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setDataElementIdentifier("26");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._458);
				}
															
				return rNIBRSError;
			}			
		};		
		return ageForNonPersonRule;		
	}

	/**
	 * 29 (Ethnicity of Victim) The Data Element associated with this error
	 * cannot be entered when Data Element 25 (Type of Victim) is not
	 * I=Individual or L=Law Enforcement Officer when Data Element 24 (Victim
	 * Connected to UCR Offense Code) contains a Crime Against Person.
	 */
	public Rule<VictimSegment> getRule458ForEthnicityOfVictim(){
		
		Rule<VictimSegment> ageForNonPersonRule = new Rule<VictimSegment>(){
									
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
			
				List<String> offenseList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasCrimeAgainstPerson = CollectionUtils.containsAny(offenseList, PERSON_OFFENSE_LIST);
								
				String sVictimType = victimSegment.getTypeOfVictim();
				
				boolean victimIsPerson = TypeOfVictimCode.I.code.equals(sVictimType)
						|| TypeOfVictimCode.L.code.equals(sVictimType);				
				
				String ethnicity = victimSegment.getEthnicity();
				
				if(hasCrimeAgainstPerson && !victimIsPerson && StringUtils.isNotEmpty(ethnicity)){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setDataElementIdentifier("29");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._458);
				}
															
				return rNIBRSError;
			}			
		};		
		return ageForNonPersonRule;		
	}

	/**
	 * (Sex of Victim) The Data Element associated with this error cannot be
	 * entered when Data Element 25 (Type of Victim) is not I=Individual or
	 * L=Law Enforcement Officer when Data Element 24 (Victim Connected to UCR
	 * Offense Code) contains a Crime Against Person.
	 */
	public Rule<VictimSegment> getRule458ForSexOfVictim(){
		
		Rule<VictimSegment> ageForNonPersonRule = new Rule<VictimSegment>(){
									
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
			
				List<String> offenseList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasCrimeAgainstPerson = CollectionUtils.containsAny(offenseList, PERSON_OFFENSE_LIST);
								
				String sVictimType = victimSegment.getTypeOfVictim();
				
				boolean victimIsPerson = TypeOfVictimCode.I.code.equals(sVictimType)
						|| TypeOfVictimCode.L.code.equals(sVictimType);				
				
				String sVictimSex = victimSegment.getSex();
				
				if(hasCrimeAgainstPerson && !victimIsPerson && StringUtils.isNotEmpty(sVictimSex)){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setDataElementIdentifier("27");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._458);
				}
															
				return rNIBRSError;
			}			
		};		
		return ageForNonPersonRule;		
	}

	/**
	 * (Race of Victim) The Data Element associated with this error cannot be
	 * entered when Data Element 25 (Type of Victim) is not I=Individual or
	 * L=Law Enforcement Officer when Data Element 24 (Victim Connected to UCR
	 * Offense Code) contains a Crime Against Person.
	 */
	public Rule<VictimSegment> getRule458ForRaceOfVictim(){
		
		Rule<VictimSegment> ageForNonPersonRule = new Rule<VictimSegment>(){
									
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
			
				List<String> offenseList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasCrimeAgainstPerson = CollectionUtils.containsAny(offenseList, PERSON_OFFENSE_LIST);
								
				String sVictimType = victimSegment.getTypeOfVictim();
				
				boolean victimIsPerson = TypeOfVictimCode.I.code.equals(sVictimType)
						|| TypeOfVictimCode.L.code.equals(sVictimType);				
				
				String victimRace = victimSegment.getRace();
				
				if(hasCrimeAgainstPerson && !victimIsPerson && StringUtils.isNotEmpty(victimRace)){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setDataElementIdentifier("28");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._458);
				}
															
				return rNIBRSError;
			}			
		};		
		return ageForNonPersonRule;		
	}

	public Rule<VictimSegment> getRule458ForResidentStatusOfVictim(){
		
		Rule<VictimSegment> ageForNonPersonRule = new Rule<VictimSegment>(){
									
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
			
				List<String> offenseList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasCrimeAgainstPerson = CollectionUtils.containsAny(offenseList, PERSON_OFFENSE_LIST);
								
				String sVictimType = victimSegment.getTypeOfVictim();
				
				boolean victimIsPerson = TypeOfVictimCode.I.code.equals(sVictimType)
						|| TypeOfVictimCode.L.code.equals(sVictimType);				
				
				String residentStatus = victimSegment.getResidentStatusOfVictim();
				
				if(hasCrimeAgainstPerson && !victimIsPerson && StringUtils.isNotEmpty(residentStatus)){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setDataElementIdentifier("30");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._458);
				}
															
				return rNIBRSError;
			}			
		};		
		return ageForNonPersonRule;		
	}

	/**
	 * (Type Injury) The Data Element associated with this error cannot be
	 * entered when Data Element 25 (Type of Victim) is not I=Individual or
	 * L=Law Enforcement Officer when Data Element 24 (Victim Connected to UCR
	 * Offense Code) contains a Crime Against Person.
	 */
	public Rule<VictimSegment> getRule458ForTypeOfInjury(){
		
		Rule<VictimSegment> ageForNonPersonRule = new Rule<VictimSegment>(){
									
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
			
				List<String> offenseList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasCrimeAgainstPerson = CollectionUtils.containsAny(offenseList, PERSON_OFFENSE_LIST);
								
				String sVictimType = victimSegment.getTypeOfVictim();
				
				boolean victimIsPerson = TypeOfVictimCode.I.code.equals(sVictimType)
						|| TypeOfVictimCode.L.code.equals(sVictimType);				
				
				List<String> typeOfInjuryList = victimSegment.getTypeOfInjuryList();
				
				if(hasCrimeAgainstPerson && !victimIsPerson && (typeOfInjuryList != null && !typeOfInjuryList.isEmpty())){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setDataElementIdentifier("33");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._458);
				}															
				return rNIBRSError;
			}			
		};		
		return ageForNonPersonRule;		
	}

	public Rule<VictimSegment> getRule458ForOffenderNumberToBeRelated(){
		
		Rule<VictimSegment> ageForNonPersonRule = new Rule<VictimSegment>(){
									
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
			
				List<String> offenseList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasCrimeAgainstPerson = CollectionUtils.containsAny(offenseList, PERSON_OFFENSE_LIST);
								
				String sVictimType = victimSegment.getTypeOfVictim();
				
				boolean victimIsPerson = TypeOfVictimCode.I.code.equals(sVictimType)
						|| TypeOfVictimCode.L.code.equals(sVictimType);				
				
				List<Integer> offenderRelatedList = victimSegment.getOffenderNumberRelatedList();
				
				if(hasCrimeAgainstPerson && !victimIsPerson && (offenderRelatedList != null && !offenderRelatedList.isEmpty())){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setDataElementIdentifier("34");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._458);
				}															
				return rNIBRSError;
			}			
		};		
		return ageForNonPersonRule;		
	}

	/**
	 * (Offender Numbers To Be Related) was entered but should only be entered
	 * if one or more of the offenses entered into Data Element 24 [Victim
	 * Connected to UCR Offense Code(s)] is a Crime Against Person or is a
	 * Robbery Offense (120). None of these types of offenses were entered.
	 */
	public Rule<VictimSegment> getRule459ForOffenderNumberToBeRelated(){
		
		Rule<VictimSegment> offenderNumPersonRule = new Rule<VictimSegment>(){
	
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;				
				
				List<Integer> offenderNumberRelatedList = victimSegment.getOffenderNumberRelatedList();
				
				boolean hasOffenderNumberRelated = offenderNumberRelatedList != null && !offenderNumberRelatedList.isEmpty();
				
				if(hasOffenderNumberRelated){
					
					List<String> offenseCodeList = victimSegment.getUcrOffenseCodeList();
					
					if(!CollectionUtils.containsAny(offenseCodeList, PERSON_OFFENSE_LIST)
							&& !offenseCodeList.contains(OffenseCode._120)){
						
						rNIBRSError = victimSegment.getErrorTemplate();
						rNIBRSError.setDataElementIdentifier("34");
						rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._459);
					}										
				}											
				return rNIBRSError;
			}			
		};		
		return offenderNumPersonRule;
	}

	/**
	 * (Relationship of Victim to Offender) Corresponding Data Element 35
	 * (Relationship of Victim to Offenders) data must be entered when Data
	 * Element 34 (Offender Numbers To Be Related) is entered with a value
	 * greater than 00.
	 */
	public Rule<VictimSegment> getRule460ForRelationshipOfVictimToOffender(){
		
		Rule<VictimSegment> relationshiOffenderRule = new Rule<VictimSegment>(){
	
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
	
				NIBRSError rNIBRSError = null;
				
				List<String> victimOffenderRelationshipList = victimSegment.getVictimOffenderRelationshipList();
								
				List<Integer> offenderNumRelatedList = victimSegment.getOffenderNumberRelatedList();
				
				boolean hasValidOffenderNumRelated = false;
				
				for(Integer iOffenderNumRelatd : offenderNumRelatedList){
				
					if(iOffenderNumRelatd != null && iOffenderNumRelatd > 0){
						
						hasValidOffenderNumRelated = true;
					}					
				}
				
				boolean hasVictimOffenderRelationship = victimOffenderRelationshipList != null 
						&& !victimOffenderRelationshipList.isEmpty();				
								
				if(hasVictimOffenderRelationship && !hasValidOffenderNumRelated){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					rNIBRSError.setDataElementIdentifier("35");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._460);
				}				
				return rNIBRSError;
			}
			
		};
		return relationshiOffenderRule;
	}

	// (Type of Victim) cannot have a value of S=Society/Public when the offense
	// is 220=Burglary/Breaking and Entering.
	public Rule<VictimSegment> getRule461ForTypeOfVictim(){
		
		Rule<VictimSegment> societyBurglaryRule = new Rule<VictimSegment>(){

			@Override
			public NIBRSError apply(VictimSegment victimSegment) {

				NIBRSError rNibrsError = null;
				
				String sVictimType = victimSegment.getTypeOfVictim();
													
				List<String> ucrOffenseCodeList = victimSegment.getUcrOffenseCodeList();
				
				if(ucrOffenseCodeList.contains(OffenseCode._220.code)
					&& TypeOfVictimCode.S.code.equals(sVictimType)){
										
					rNibrsError = victimSegment.getErrorTemplate();
					
					rNibrsError.setNIBRSErrorCode(NIBRSErrorCode._461);
					rNibrsError.setDataElementIdentifier("25");					
				}												
				return rNibrsError;
			}};
		
		return societyBurglaryRule;
	}
	

	
	/**
	 * (Type of Victim) contains a Crime Against Person, but Data Element 25
	 * (Type of Victim) is not I=Individual or L=Law Enforcement Officer when
	 * Data Element 24 (Victim Connected to UCR Offense Code) contains a Crime
	 * Against Person.
	 */
	public Rule<VictimSegment> getRule464ForTypeOfVictim(){
		
		Rule<VictimSegment> personTypeRule = new Rule<VictimSegment>(){
									
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
			
				List<String> offenseList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasCrimeAgainstPerson = CollectionUtils.containsAny(offenseList, PERSON_OFFENSE_LIST);
								
				String sVictimType = victimSegment.getTypeOfVictim();
				
				boolean victimIsPerson = TypeOfVictimCode.I.code.equals(sVictimType)
						|| TypeOfVictimCode.L.code.equals(sVictimType);
				
				if(hasCrimeAgainstPerson && !victimIsPerson){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					
					rNIBRSError.setDataElementIdentifier("25");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._464);
				}
												
				return rNIBRSError;
			}			
		};		
		return personTypeRule;		
	}
	
	
	
	
	/**
	 * (Type of Victim) contains a Crime Against Society, but Data Element 25 (Type of Victim) is not S=Society.
	 */
	public Rule<VictimSegment> geRule465ForTypeOfVictim(){
				
		Rule<VictimSegment> crimeSocietyRule = new Rule<VictimSegment>(){
			
			List<String> crimeAgainstSocietyOffenseList = Arrays.asList(OffenseCode._720.code,
					OffenseCode._35A.code, OffenseCode._35B.code, OffenseCode._39A.code, OffenseCode._39B.code,
					OffenseCode._39C.code, OffenseCode._39D.code, OffenseCode._370.code, OffenseCode._40A.code,
					OffenseCode._40B.code, OffenseCode._40C.code, OffenseCode._520.code);			

			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNibrsError = null;
				
				List<String> offenseCodeList = victimSegment.getUcrOffenseCodeList();
								
				boolean hasCrimeAgainstSociety = CollectionUtils.containsAny(offenseCodeList, crimeAgainstSocietyOffenseList); 
				
				String sVictimType = victimSegment.getTypeOfVictim();
				
				if(hasCrimeAgainstSociety 
					&& !TypeOfVictimCode.S.code.equals(sVictimType)){
					
					rNibrsError = victimSegment.getErrorTemplate();
					
					rNibrsError.setNIBRSErrorCode(NIBRSErrorCode._465);
					rNibrsError.setDataElementIdentifier("25");					
				}				
				return rNibrsError;
			}					
		};						
		return crimeSocietyRule;
	}
	


	/**
	 * (Type of Victim) contains a Crime Against Property, but Data Element 25
	 * (Type of Victim) is S=Society. This is not an allowable code for Crime
	 * Against Property offenses.
	 */
	public Rule<VictimSegment> getRule467ForTypeOfVictim(){
		
		Rule<VictimSegment> typeVictimPropertyOffenseRule = new Rule<VictimSegment>(){
				
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {

				NIBRSError rNIBRSError = null;
				
				List<String> offenseCodeList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasCrimeAgainstProperty = CollectionUtils.containsAny(offenseCodeList, PROPERTY_CRIM_LIST);
				
				String sVictimType = victimSegment.getTypeOfVictim();
				
				if(TypeOfVictimCode.S.equals(sVictimType) 
						&& hasCrimeAgainstProperty){
					
					rNIBRSError = victimSegment.getErrorTemplate();
					
					rNIBRSError.setDataElementIdentifier("25");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._467);
				}								
				return rNIBRSError;
			}
		};			
		return typeVictimPropertyOffenseRule;
	}
	
	
	
	
	
	
	// TODO 483
	
	/**
	 * (Relationship of Victim to Offender) cannot be entered when Data Element
	 * 34 (Offender Number to be Related) is zero. Zero means that the number of
	 * offenders is unknown; therefore, the relationship cannot be entered.
	 */
	public Rule<VictimSegment> getRule468ForRelationshipOfVictimToOffender(){
		
		Rule<VictimSegment> unknownOffendersRule = new Rule<VictimSegment>(){
	
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
					
				List<String> victimOffenderRelationshipList = victimSegment.getVictimOffenderRelationshipList();
								
				List<Integer> offenderNumberRelatedList = victimSegment.getOffenderNumberRelatedList();
				
				if(offenderNumberRelatedList != null && offenderNumberRelatedList.contains(0)){
					if(victimOffenderRelationshipList != null && !victimOffenderRelationshipList.isEmpty()){
						
						rNIBRSError = victimSegment.getErrorTemplate();
						rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._468);
						rNIBRSError.setDataElementIdentifier("35");						
					}					
				}								
				return rNIBRSError;
			}			
		};
		return unknownOffendersRule; 		
	}

	public Rule<VictimSegment> getRule469ForSexOfVictim(){
		
		Rule<VictimSegment> sexOfVictimForStatutoryRapeRule = new Rule<VictimSegment>(){

			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
								
				List<String> offenseList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasRapeOffense = offenseList.contains(OffenseCode._36B)
						|| offenseList.contains(OffenseCode._11A);				
				
				String victimSexCode = victimSegment.getSex();
				
				boolean isMaleOrFemale = SexOfVictimCode.F.code.equals(victimSexCode)
						|| SexOfVictimCode.M.code.equals(victimSexCode);
																
				if(hasRapeOffense && !isMaleOrFemale){
					
					rNIBRSError = victimSegment.getErrorTemplate();					
					rNIBRSError.setDataElementIdentifier("27");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._469);
				}				
				
				return rNIBRSError;
			}						
		};		
		return sexOfVictimForStatutoryRapeRule;
	}
	
	
	
	public Rule<VictimSegment> getRule481ForAgeOfVictim(){
		
		Rule<VictimSegment> ageOfVictimUnder18ForRapeRule = new Rule<VictimSegment>(){
	
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
								
				List<String> offenseList = victimSegment.getUcrOffenseCodeList();
				
				boolean hasStatRape = offenseList.contains(OffenseCode._36B);				
				
				NIBRSAge age = victimSegment.getAge();
				
				boolean ageLessThan18 = age != null && age.getAgeMax() < 18;
				
				if(hasStatRape && !ageLessThan18){
					
					rNIBRSError = victimSegment.getErrorTemplate();					
					rNIBRSError.setDataElementIdentifier("26");
					rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._481);
				}				
				
				return rNIBRSError;
			}						
		};		
		return ageOfVictimUnder18ForRapeRule;
	}

	
	/**
	 * (Type of Victim) cannot be L=Law Enforcement Officer unless Data Element
	 * 24 (Victim Connected to UCR Offense Code) is one of the following:
	 * 09A=Murder & Non-negligent Manslaughter 13A=Aggravated Assault 13B=Simple
	 * Assault 13C=Intimidation
	 */
	public Rule<VictimSegment> getRule482ForTypeOfVictim(){
		
		Rule<VictimSegment> typeVictimLawOfficerRule = new Rule<VictimSegment>(){
	
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
				
				NIBRSError rNIBRSError = null;
				
				String sVictimType = victimSegment.getTypeOfVictim();
				
				boolean isLawOfficerVictimType = TypeOfVictimCode.L.equals(sVictimType);
				
				if(isLawOfficerVictimType){
					
					List<String> lawOfficerVictimOffenseList = Arrays.asList(
							OffenseCode._09A.code, OffenseCode._13A.code,
							OffenseCode._13B.code, OffenseCode._13C.code);
										
					List<String> victimOffenseList = victimSegment.getUcrOffenseCodeList();		
										
					if(!lawOfficerVictimOffenseList.containsAll(victimOffenseList)){
	
						rNIBRSError = victimSegment.getErrorTemplate();
						
						rNIBRSError.setDataElementIdentifier("25");
						rNIBRSError.setNIBRSErrorCode(NIBRSErrorCode._482);
					};															
				}				
				return rNIBRSError;
			}			
		};		
		return typeVictimLawOfficerRule;
	}
	

	/**
	 * Type of Officer Activity/Circumstance) Data Element 25B (Officer
	 * Assignment Type), Data Element 25C (Officer–ORI Other Jurisdiction), Data
	 * Element 26 (Age of Victim), Data Element 27 (Sex of Victim), Data Element
	 * 28 (Race of Victim), Data Element 29 (Ethnicity of Victim), Data Element
	 * 30 (Resident Status of Victim), and Data Element 34 (Offender Number to
	 * be Related) can only be entered when Data Element 25 (Type of Victim) is
	 * I=Individual or L=Law Enforcement Officer.
	 */
	public Rule<VictimSegment> getRule483ForTypeOfOfficerActivity(){
		
		Rule<VictimSegment> personReqFieldsRule = new Rule<VictimSegment>(){
	
			NIBRSError nibrsError = null;
			
			@Override
			public NIBRSError apply(VictimSegment victimSegment) {
	
				String assignmentType = victimSegment.getOfficerAssignmentType();
				
				String ori = victimSegment.getOfficerOtherJurisdictionORI();
				
				NIBRSAge age = victimSegment.getAge();
				
				String race = victimSegment.getRace();
				
				String sex = victimSegment.getSex();
				
				String ethnicity = victimSegment.getEthnicity();
				
				String residentStatus = victimSegment.getResidentStatusOfVictim();				
				
				List<Integer> relatedOffenderList = victimSegment.getOffenderNumberRelatedList();
								
				String typeOfVictim = victimSegment.getTypeOfVictim();
				
				boolean isPersonVictim = TypeOfVictimCode.I.code.equals(typeOfVictim) 
						|| TypeOfVictimCode.L.code.equals(typeOfVictim); 
								
				if(StringUtils.isNotEmpty(assignmentType)
					|| StringUtils.isNotEmpty(ori)
					|| age != null
					|| StringUtils.isNotEmpty(race)
					|| StringUtils.isNotEmpty(sex)
					|| StringUtils.isNotEmpty(ethnicity)
					|| StringUtils.isNotEmpty(residentStatus)
					|| (relatedOffenderList != null && !relatedOffenderList.isEmpty()) 
					&& !isPersonVictim){
					
					nibrsError = victimSegment.getErrorTemplate();
					nibrsError.setDataElementIdentifier("25B");
					nibrsError.setNIBRSErrorCode(NIBRSErrorCode._483);
				}								
				return nibrsError;
			}			
		};		
		return personReqFieldsRule;
	}
	
}
