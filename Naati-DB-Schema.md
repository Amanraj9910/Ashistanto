# MyNaati UI Database Schema
**Database:** `Local_Naati` (Microsoft SQL Server)  
**Schema:** `dbo` (primary), `HangFire`, `MyNaatiHangFire`  
**Total Tables:** 348 (277 core + 71 backup/system)  
**Total Foreign Keys:** 482  
**Views:** 18 | **Stored Procedures:** 98 | **Indexes:** 313
> [!NOTE]
> This document covers only the **277 core tables** — backup tables (suffix `_Backup_*`) and legacy ASP.NET membership tables (`aspnet_*`) are excluded. Tables are organized by functional domain.
---
## Table of Contents
1. [Entity & Person Management](#1-entity--person-management)
2. [Credentials & Certifications](#2-credentials--certifications)
3. [Credential Applications & Workflows](#3-credential-applications--workflows)
4. [Testing & Examination](#4-testing--examination)
5. [Test Materials & Resources](#5-test-materials--resources)
6. [Panel & Examiner Management](#6-panel--examiner-management)
7. [Professional Development](#7-professional-development)
8. [Email & Communication](#8-email--communication)
9. [Orders, Payments & Finance](#9-orders-payments--finance)
10. [File Storage](#10-file-storage)
11. [Security & Access Control](#11-security--access-control)
12. [System Configuration](#12-system-configuration)
13. [Location & Geography](#13-location--geography)
14. [Role Players](#14-role-players)
15. [Lookup Tables (tlu*)](#15-lookup-tables-tlu)
16. [HangFire Job Queue](#16-hangfire-job-queue)
17. [Miscellaneous](#17-miscellaneous)
18. [Database Views](#18-database-views)
---
## 1. Entity & Person Management
Core entities representing people, institutions, and their associated data.
### tblEntity
The root entity table — all persons and institutions link here.
| Column | Type | Nullable | PK | Default |
|--------|------|----------|-----|---------|
| EntityId | int | NO | ✅ | |
| NaatiNumber | int | YES | | |
| EntityTypeId | int | NO | | |
| ODHidden | bit | NO | | |
| ModifiedByNaati | bit | NO | | |
| ModifiedDate | datetime | NO | | |
| ModifiedUser | int | NO | | |
**FKs:** None outbound (this is a root table)
---
### tblPerson
Individuals registered in the NAATI system.
| Column | Type | Length | Nullable | PK | Default |
|--------|------|--------|----------|-----|---------|
| PersonId | int | | NO | ✅ | |
| EntityId | int | | NO | | |
| DateOfBirth | datetime | | YES | | |
| GenderId | int | | NO | | |
| BirthCountryId | int | | YES | | |
| SponsorInstitutionId | int | | YES | | |
| Deceased | bit | | NO | | |
| LegislationExempt | bit | | NO | | |
| AccessDisabled | bit | | NO | | (0) |
| AccessDisabledTypeId | tinyint | | YES | | |
| EPortalAccessDisabledExpiryDate | datetime | | YES | | |
| ODHidden | bit | | NO | | |
| ODRestricted | bit | | NO | | |
| ModifiedByNaati | bit | | NO | | |
| ModifiedDate | datetime | | NO | | |
| ModifiedUser | int | | NO | | |
**FKs:**
- `EntityId` → `tblEntity.EntityId`
- `BirthCountryId` → `tblCountry.CountryId`
- `SponsorInstitutionId` → `tblInstitution.InstitutionId`
---
### tblPersonName
Name history for persons (supports name changes with effective dates).
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| PersonNameId | int | | NO | ✅ |
| PersonId | int | | NO | |
| TitleId | int | | YES | |
| GivenName | varchar | 100 | NO | |
| MiddleName | varchar | 100 | YES | |
| Surname | varchar | 100 | NO | |
| EffectiveDate | datetime | | NO | |
| DisplayOnOd | bit | | NO | |
| ModifiedByNaati | bit | | NO | |
| ModifiedDate | datetime | | NO | |
| ModifiedUser | int | | NO | |
**FKs:** `PersonId` → `tblPerson`, `TitleId` → `tluTitle`
---
### tblPersonImage
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| PersonImageId | int | NO | ✅ |
| PersonId | int | NO | |
| Data | varbinary | NO | |
| ContentType | varchar(50) | NO | |
| Active | bit | NO | |
| UploadedDate | datetime | NO | |
**FKs:** `PersonId` → `tblPerson`
---
### tblPersonAttachment
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| PersonAttachmentId | int | NO | ✅ |
| PersonId | int | NO | |
| StoredFileId | int | NO | |
| Description | nvarchar(255) | YES | |
| Deleted | bit | NO | |
**FKs:** `PersonId` → `tblPerson`, `StoredFileId` → `tblStoredFile`
---
### tblPersonAccessBlock
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| PersonAccessBlockId | int | NO | ✅ |
| PersonId | int | NO | |
| UserId | int | NO | |
| AccessDisabledTypeId | tinyint | NO | |
**FKs:** `PersonId` → `tblPerson`, `UserId` → `tblUser`, `AccessDisabledTypeId` → `tluAccessDisabledType`
---
### tblInstitution
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| InstitutionId | int | | NO | ✅ |
| EntityId | int | | NO | |
| Active | bit | | NO | |
| Website | varchar | 200 | YES | |
| ODHidden | bit | | NO | |
| ModifiedByNaati | bit | | NO | |
| ModifiedDate | datetime | | NO | |
| ModifiedUser | int | | NO | |
**FKs:** `EntityId` → `tblEntity`
---
### tblInstitutionName
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| InstitutionNameId | int | | NO | ✅ |
| InstitutionId | int | | NO | |
| Name | varchar | 200 | NO | |
| EffectiveDate | datetime | | YES | |
| ModifiedByNaati | bit | | NO | |
| ModifiedDate | datetime | | NO | |
| ModifiedUser | int | | NO | |
**FKs:** `InstitutionId` → `tblInstitution`
---
### tblAddress
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| AddressId | int | | NO | ✅ |
| EntityId | int | | NO | |
| PostcodeId | int | | YES | |
| AddressTypeId | int | | NO | |
| ODAddressVisibilityTypeId | int | | YES | |
| AddressLine1 | nvarchar | 150 | YES | |
| AddressLine2 | nvarchar | 150 | YES | |
| AddressLine3 | nvarchar | 150 | YES | |
| AddressLine4 | nvarchar | 150 | YES | |
| Active | bit | | NO | |
| FreeFormPostcode | nvarchar | 10 | YES | |
| FreeFormSuburb | nvarchar | 100 | YES | |
| FreeFormState | nvarchar | 100 | YES | |
| FreeFormCountry | nvarchar | 100 | YES | |
**FKs:** `EntityId` → `tblEntity`, `PostcodeId` → `tblPostcode`, `ODAddressVisibilityTypeId` → `tblODAddressVisibilityType`
---
### tblPhone
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| PhoneId | int | | NO | ✅ |
| EntityId | int | | NO | |
| PhoneTypeId | int | | NO | |
| Number | varchar | 20 | NO | |
| Active | bit | | NO | |
**FKs:** `EntityId` → `tblEntity`
---
### tblEmail
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| EmailId | int | | NO | ✅ |
| EntityId | int | | NO | |
| EmailTypeId | int | | NO | |
| Address | nvarchar | 200 | NO | |
| Active | bit | | NO | |
**FKs:** `EntityId` → `tblEntity`
---
### tblContactPerson
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| ContactPersonId | int | NO | ✅ |
| InstitutionId | int | NO | |
| GivenName | varchar(100) | NO | |
| Surname | varchar(100) | NO | |
| Position | varchar(100) | YES | |
**FKs:** `InstitutionId` → `tblInstitution`
---
### tblEntityNote
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| EntityNoteId | int | NO | ✅ |
| EntityId | int | NO | |
| NoteId | int | NO | |
**FKs:** `EntityId` → `tblEntity`, `NoteId` → `tblNote`
---
### tblAboriginalTorresIdentity / tblAboriginalTorresOption
Aboriginal and Torres Strait Islander identity tracking.
**tblAboriginalTorresIdentity** — `AboriginalTorresIdentityId` (PK), `EntityId` → `tblEntity`, `AboriginalTorresOptionId` → `tblAboriginalTorresOption`
**tblAboriginalTorresOption** — `AboriginalTorresOptionId` (PK), [Name](file:///d:/naati-productionsourcecode/F1Solutions.Naati.Common.Dal.NHibernate/Configuration/NHibernateSession.cs#369-375), `ModifiedUser` → `tblUser`
---
### tblDialects / tblExperience / tblService / tblQualification
Entity-linked supplementary data tables:
- **tblDialects** — `DialectsId` (PK), `EntityId` → `tblEntity`
- **tblExperience** — `ExperienceId` (PK), `EntityId` → `tblEntity`
- **tblService** — `ServiceId` (PK), `EntityId` → `tblEntity`
- **tblQualification** — `QualificationId` (PK), `EntityId` → `tblEntity`
---
## 2. Credentials & Certifications
### tblCredentialType
Defines types of credentials (e.g., Certified Translator, Certified Interpreter).
| Column | Type | Length | Nullable | PK | Default |
|--------|------|--------|----------|-----|---------|
| CredentialTypeId | int | | NO | ✅ | |
| Name | nvarchar | 50 | NO | | |
| CredentialCategoryId | int | | NO | | |
| SkillTypeId | int | | NO | | |
| DisplayName | nvarchar | 100 | NO | | |
| Active | bit | | NO | | (1) |
| CanUpgrade | bit | | NO | | (0) |
| CanDowngrade | bit | | NO | | (0) |
| CanCrossSkill | bit | | NO | | (0) |
| OdDisplayAllLanguages | bit | | NO | | (0) |
| ModifiedByNaati | bit | | NO | | |
| ModifiedDate | datetime | | NO | | |
| ModifiedUser | int | | NO | | |
**FKs:** `CredentialCategoryId` → `tblCredentialCategory`, `SkillTypeId` → `tblSkillType`, `ModifiedUser` → `tblUser`
---
### tblCredential
Issued credentials for persons.
| Column | Type | Length | Nullable | PK | Default |
|--------|------|--------|----------|-----|---------|
| CredentialId | int | | NO | ✅ | |
| PersonId | int | | NO | | |
| CertificationPeriodId | int | | YES | | |
| CredentialTypeId | int | | NO | | |
| SkillId | int | | NO | | |
| CredentialStatusTypeId | int | | NO | | |
| CredentialNumber | varchar | 50 | YES | | |
| IssueDate | datetime | | YES | | |
| SuspendedDate | datetime | | YES | | |
| ExpiryDate | datetime | | YES | | |
| DisplayOnOd | bit | | NO | | (1) |
| ModifiedByNaati | bit | | NO | | |
| ModifiedDate | datetime | | NO | | |
| ModifiedUser | int | | NO | | |
| LegislationExempt | bit | | NO | | (0) |
**FKs:** `CertificationPeriodId` → `tblCertificationPeriod`
---
### tblCredentialCategory
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| CredentialCategoryId | int | NO | ✅ |
| Name | varchar(50) | NO | |
| TlaCategoryId | int | YES | |
| ModifiedUser | int | NO | |
**FKs:** `ModifiedUser` → `tblUser`
---
### tblCredentialStatusType
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| CredentialStatusTypeId | int | NO | ✅ |
| Name | varchar(50) | NO | |
| DisplayName | nvarchar(50) | NO | |
| ModifiedUser | int | NO | |
---
### tblCredentialRequest
Requests for new/renewed credentials.
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| CredentialRequestId | int | | NO | ✅ |
| CredentialApplicationId | int | | NO | |
| CredentialTypeId | int | | NO | |
| SkillId | int | | NO | |
| CredentialRequestPathTypeId | int | | NO | |
| CredentialRequestStatusTypeId | int | | NO | |
| StatusChangeUserId | int | | NO | |
| StatusChangeDate | datetime | | NO | |
**FKs:**
- `CredentialApplicationId` → `tblCredentialApplication`
- `CredentialTypeId` → `tblCredentialType`
- `SkillId` → `tblSkill`
- `CredentialRequestPathTypeId` → `tblCredentialRequestPathType`
- `CredentialRequestStatusTypeId` → `tblCredentialRequestStatusType`
- `StatusChangeUserId` → `tblUser`
---
### tblCredentialCredentialRequest
Links credentials to their originating requests.
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| CredentialCredentialRequestId | int | NO | ✅ |
| CredentialId | int | NO | |
| CredentialRequestId | int | NO | |
**FKs:** `CredentialId` → `tblCredential`, `CredentialRequestId` → `tblCredentialRequest`
---
### tblCertificationPeriod
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| CertificationPeriodId | int | NO | ✅ |
| PersonId | int | NO | |
| StartDate | datetime | NO | |
| ExpiryDate | datetime | NO | |
**FKs:** `PersonId` → `tblPerson`
---
### tblCredentialAttachment / tblCredentialQrCode / tblCredentialFeeProduct
Supporting credential tables:
- **tblCredentialAttachment** — links `CredentialId` → `tblCredential`, `StoredFileId` → `tblStoredFile`
- **tblCredentialQrCode** — `CredentialId` → `tblCredential`
- **tblCredentialFeeProduct** — links credentials to products/fees via `CredentialTypeId`, `CredentialApplicationTypeId`, `ProductSpecificationId`, `CredentialApplicationRefundPolicyId`
---
### Credential Type Paths & Cross-Skills
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblCredentialTypeCrossSkill | Cross-skill paths | CredentialTypeFromId/ToId → tblCredentialType |
| tblCredentialTypeDowngradePath | Downgrade paths | CredentialTypeFromId/ToId → tblCredentialType |
| tblCredentialTypeUpgradePath | Upgrade paths | CredentialTypeFromId/ToId → tblCredentialType |
| tblCredentialTypeTemplate | Templates per type | CredentialTypeId → tblCredentialType, StoredFileId → tblStoredFile |
| tblCredentialTypeTestMaterialDomain | Material domains | CredentialTypeId → tblCredentialType, TestMaterialDomainId → tblTestMaterialDomain |
| tblCredentialPrerequisite | Prerequisites | CredentialTypeId, CredentialApplicationTypeId |
| tblCredentialPrerequisiteExemption | Exemptions | CredentialTypeId, PersonId, SkillId |
---
## 3. Credential Applications & Workflows
### tblCredentialApplication
Central application entity.
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| CredentialApplicationId | int | | NO | ✅ |
| CredentialApplicationTypeId | int | | NO | |
| CredentialApplicationStatusTypeId | int | | NO | |
| PersonId | int | | NO | |
| SponsorInstitutionId | int | | YES | |
| ReceivingOfficeId | int | | NO | |
| EnteredUserId | int | | NO | |
| StatusChangeUserId | int | | NO | |
| OwnedByUserId | int | | NO | |
| ApplicationDate | datetime | | NO | |
| StatusChangeDate | datetime | | NO | |
| ApplicationNumber | varchar | 50 | YES | |
| SubmittedByApplicant | bit | | NO | |
| EthicsAndConduct | bit | | NO | |
| PrivacyAgreement | bit | | NO | |
**FKs:**
- `CredentialApplicationTypeId` → `tblCredentialApplicationType`
- `CredentialApplicationStatusTypeId` → `tblCredentialApplicationStatusType`
- `PersonId` → `tblPerson`
- `SponsorInstitutionId` → `tblInstitution`
- `ReceivingOfficeId` → `tblOffice`
- `EnteredUserId` / `StatusChangeUserId` / `OwnedByUserId` → `tblUser`
---
### tblCredentialApplicationType
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| CredentialApplicationTypeId | int | | NO | ✅ |
| Name | varchar | 100 | NO | |
| CredentialApplicationTypeCategoryId | int | | NO | |
| DisplayName | nvarchar | 100 | NO | |
| Active | bit | | NO | |
**FKs:** `CredentialApplicationTypeCategoryId` → `tblCredentialApplicationTypeCategory`
---
### Application Form System
A dynamic forms engine for credential applications:
| Table | Purpose |
|-------|---------|
| tblCredentialApplicationForm | Form definitions linked to application types |
| tblCredentialApplicationFormSection | Form sections |
| tblCredentialApplicationFormQuestion | Questions within sections |
| tblCredentialApplicationFormQuestionType | Question types with answer types |
| tblCredentialApplicationFormAnswerType | Answer format definitions |
| tblCredentialApplicationFormAnswerOption | Answer options per question type |
| tblCredentialApplicationFormQuestionAnswerOption | Links questions to specific answer options |
| tblCredentialApplicationFormQuestionLogic | Conditional logic (hide/show based on answers, credential type, skill) |
| tblCredentialApplicationFormActionType | Actions triggered by answer selections |
| tblCredentialApplicationFormAnswerOptionActionType | Links answer options to action types |
| tblCredentialApplicationFormAnswerOptionDocumentType | Document types required for specific answers |
| tblCredentialApplicationFormCredentialType | Links forms to credential types |
| tblCredentialApplicationFormUserType | Form user type definitions |
---
### Application Data & Attachments
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblCredentialApplicationFieldData | Field values for applications | CredentialApplicationId, CredentialApplicationFieldId |
| tblCredentialApplicationAttachment | Attached files | CredentialApplicationId, StoredFileId |
| tblCredentialApplicationEmailMessage | Email messages | CredentialApplicationId, EmailMessageId |
| tblCredentialApplicationNote | Notes | CredentialApplicationId, NoteId |
| tblCredentialApplicationRefund | Refund records | CredentialWorkflowFeeId, RefundMethodTypeId, UserId |
---
### Credential Application Fields
| Table | Purpose |
|-------|---------|
| tblCredentialApplicationField | Field definitions per application type |
| tblCredentialApplicationFieldCategory | Field groupings |
| tblCredentialApplicationFieldOption | Option lists for fields |
| tblCredentialApplicationFieldOptionOption | Individual options within option lists |
---
### Credential Request Supporting Tables
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblCredentialRequestCredentialRequest | Links related requests | OriginalCredentialRequestId, AssociatedCredentialRequestId |
| tblCredentialRequestAssociationType | Association types | ModifiedUser → tblUser |
| tblCredentialRequestFieldData | Field data per request | CredentialRequestId, CredentialApplicationFieldId |
| tblCredentialRequestPathType | Path types (new, renewal, etc.) | |
| tblCredentialRequestStatusType | Status types | |
---
### Workflow & Fees
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblCredentialWorkflowFee | Fee items per application | CredentialApplicationId, CredentialRequestId, ProductSpecificationId |
| tblCredentialWorkflowActionEmailTemplate | Email actions in workflows | CredentialApplicationTypeId, SystemActionEmailTemplateId |
| tblCredentialApplicationRefundPolicy | Refund policies | |
| tblRefundPolicyParameter | Policy parameters | CredentialApplicationRefundPolicyId |
---
## 4. Testing & Examination
### tblTestSession
Scheduled test sessions at venues.
| Column | Type | Length | Nullable | PK | Default |
|--------|------|--------|----------|-----|---------|
| TestSessionId | int | | NO | ✅ | |
| VenueId | int | | NO | | |
| Name | nvarchar | 100 | NO | | |
| TestDateTime | datetime | | NO | | |
| ArrivalTime | int | | YES | | |
| Duration | int | | YES | | |
| CredentialTypeId | int | | NO | | |
| Completed | bit | | NO | | |
| PublicNote | nvarchar | 1000 | YES | | |
| AllowSelfAssign | bit | | NO | | (0) |
| OverrideVenueCapacity | bit | | NO | | (0) |
| Capacity | int | | YES | | |
| RehearsalDateTime | datetime | | YES | | |
| RehearsalNotes | nvarchar | max | YES | | |
| DefaultTestSpecificationId | int | | NO | | |
| AllowAvailabilityNotice | bit | | NO | | |
| NewCandidatesOnly | bit | | YES | | |
| IsActive | bit | | NO | | (1) |
**FKs:** `VenueId` → `tblVenue`, `CredentialTypeId` → `tblCredentialType`, `DefaultTestSpecificationId` → `tblTestSpecification`
---
### tblTestSitting
Individual candidate test sittings within sessions.
| Column | Type | Nullable | PK | Default |
|--------|------|----------|-----|---------|
| TestSittingId | int | NO | ✅ | |
| TestSessionId | int | NO | | |
| CredentialRequestId | int | NO | | |
| Rejected | bit | NO | | (0) |
| Sat | bit | NO | | (0) |
| Supplementary | bit | NO | | (0) |
| TestSpecificationId | int | NO | | |
| AllocatedDate | datetime | NO | | |
| RejectedDate | datetime | YES | | |
**FKs:** `TestSessionId` → `tblTestSession`, `CredentialRequestId` → `tblCredentialRequest`, `TestSpecificationId` → `tblTestSpecification`
---
### tblTestResult
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| TestResultId | int | | NO | ✅ |
| TestSittingId | int | | NO | |
| CurrentJobId | int | | YES | |
| ResultTypeId | int | | NO | |
| CommentsGeneral | varchar | 500 | NO | |
| ThirdExaminerRequired | bit | | NO | |
| ReviewInvoiceLineId | int | | YES | |
| ProcessedDate | datetime | | YES | |
| SatDate | datetime | | YES | |
| ResultChecked | bit | | NO | |
| AllowCalculate | bit | | NO | |
| IncludePreviousMarks | bit | | YES | |
| EligibleForConcededPass | bit | | NO | |
| EligibleForSupplementary | bit | | NO | |
| AllowIssue | bit | | NO | |
| AutomaticIssuingExaminer | bit | | YES | |
**FKs:** `TestSittingId` → `tblTestSitting`, `CurrentJobId` → `tblJob`, `ResultTypeId` → `tluResultType`
---
### tblTestSpecification
Defines test structure and rules.
| Column | Type | Length | Nullable | PK | Default |
|--------|------|--------|----------|-----|---------|
| TestSpecificationId | int | | NO | ✅ | |
| Description | varchar | 500 | NO | | |
| CredentialTypeId | int | | NO | | |
| Active | bit | | NO | | (1) |
| ResultAutoCalculation | bit | | NO | | (0) |
| AutomaticIssuing | bit | | NO | | |
| MaxScoreDifference | float | | YES | | |
| TestMaterialReminder | bit | | NO | | (0) |
| TestMaterialReminderDays | int | | NO | | (0) |
**FKs:** `CredentialTypeId` → `tblCredentialType`
---
### Test Component System
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblTestComponent | Individual test components | TestSpecificationId, TypeId → tblTestComponentType |
| tblTestComponentType | Component type definitions | TestSpecificationId, TestComponentBaseTypeId |
| tblTestComponentBaseType | Base types (e.g., Written, Oral) | |
| tblTestComponentResult | Results per component | TestResultId → tblTestResult, MarkingResultTypeId |
| tblTestComponentTypeStandardMarkingScheme | Standard marks per component type | TestComponentTypeId |
| tblTestSpecificationTestComponentType | Links specs to component types | TestSpecificationId, TestComponentTypeId |
| tblTestSpecificationStandardMarkingScheme | Overall pass marks | TestSpecificationId |
| tblTestSpecificationAttachment | Attachments for specs | TestSpecificationId, StoredFileId |
---
### Test Sitting Supporting Tables
| Table | Key FKs |
|-------|---------|
| tblTestSittingDocument | TestSittingId, StoredFileId |
| tblTestSittingNote | TestSittingId, NoteId |
| tblTestSittingTestMaterial | TestSittingId, TestMaterialId, TestComponentId |
---
### Test Session Skills & Results
| Table | Key FKs |
|-------|---------|
| tblTestSessionSkill | TestSessionId, SkillId |
| tblTestResultEligibilityType | (lookup — PK: TestResultEligibilityTypeId) |
| tblTestResultRubricTestComponentResult | TestResultId, RubricTestComponentResultId |
| tblTestStatusType | (lookup — PK: TestStatusTypeId) |
---
### Rubric Marking System
Complex rubric-based assessment:
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblRubricMarkingCompetency | Competency definitions | TestComponentTypeId |
| tblRubricMarkingAssessmentCriterion | Assessment criteria | RubricMarkingCompetencyId |
| tblRubricMarkingBand | Marking bands per criterion | RubricMarkingAssessmentCriterionId |
| tblRubricTestComponentResult | Results per component | TestComponentId, MarkingResultTypeId |
| tblRubricAssessementCriterionResult | Results per criterion | RubricMarkingAssessmentCriterionId, RubricMarkingBandId, RubricTestComponentResultId |
| tblRubricQuestionPassRule | Pass rules | TestSpecificationId, RubricMarkingAssessmentCriterionId, TestComponentTypeId |
| tblRubricTestBandRule | Band rules | TestSpecificationId, RubricMarkingAssessmentCriterionId, TestResultEligibilityTypeId |
| tblRubricTestQuestionRule | Question rules | TestSpecificationId, TestComponentTypeId, TestResultEligibilityTypeId |
---
## 5. Test Materials & Resources
### tblTestMaterial
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| TestMaterialId | int | | NO | ✅ |
| Description | varchar | 500 | NO | |
| TestMaterialTypeId | int | | NO | |
| SkillId | int | | NO | |
| LanguageId | int | | YES | |
| TestComponentTypeId | int | | YES | |
| TestMaterialDomainId | int | | NO | |
| Active | bit | | NO | |
| Quarantined | bit | | NO | |
| UsageCount | int | | NO | |
| ShowExaminers | bit | | NO | |
**FKs:** `SkillId` → `tblSkill`, `LanguageId` → `tblLanguage`, `TestComponentTypeId` → `tblTestComponentType`, `TestMaterialDomainId` → `tblTestMaterialDomain`, `TestMaterialTypeId` → `tblTestMaterialType`
---
### Material Request System
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblMaterialRequest | Material creation requests | PanelId, ProductSpecificationId, SourceMaterialId, OutputMaterialId |
| tblMaterialRequestRound | Review rounds | MaterialRequestId, MaterialRequestRoundStatusTypeId |
| tblMaterialRequestRoundAttachment | Round attachments | MaterialRequestRoundId, StoredFileId |
| tblMaterialRequestRoundLink | Round links to persons | MaterialRequestRoundId, PersonId, UserId |
| tblMaterialRequestPanelMembership | Panel member assignments | MaterialRequestId, PanelMembershipId |
| tblMaterialRequestPanelMembershipTask | Tasks for panel members | MaterialRequestPanelMembershipId, MaterialRequestTaskTypeId |
| tblMaterialRequestPayroll | Payment tracking | MaterialRequestPanelMembershipId |
| tblMaterialRequestNote | Notes | MaterialRequestId, NoteId |
| tblMaterialRequestPublicNote | Public notes | MaterialRequestId, NoteId |
| tblMaterialRequestEmailMessage | Emails | MaterialRequestId, EmailMessageId |
| tblMaterialRequestStatusType | Status types | |
| tblMaterialRequestRoundStatusType | Round status types | |
| tblMaterialRequestPanelMembershipType | Membership types | |
| tblMaterialRequestTaskType | Task types | |
---
### Supporting Material Tables
| Table | Key FKs |
|-------|---------|
| tblTestMaterialAttachment | TestMaterialId, StoredFileId |
| tblTestMaterialLink | FromTestMaterialId, ToTestMaterialId, TestMaterialLinkTypeId |
| tblTestMaterialDomain | (lookup) |
| tblTestMaterialType | (lookup) |
| tblTestMaterialLinkType | (lookup) |
---
## 6. Panel & Examiner Management
### tblPanel
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| PanelId | int | NO | ✅ |
| Name | varchar(100) | NO | |
| LanguageId | int | YES | |
| PanelTypeId | int | NO | |
| Active | bit | NO | |
**FKs:** `LanguageId` → `tblLanguage`, `PanelTypeId` → `tluPanelType`
---
### tblPanelMembership
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| PanelMembershipId | int | NO | ✅ |
| PanelId | int | NO | |
| PersonId | int | NO | |
| PanelRoleId | int | NO | |
| Active | bit | NO | |
| CoordinatorCredentialTypes | bit | NO | |
**FKs:** `PanelId` → `tblPanel`, `PersonId` → `tblPerson`, `PanelRoleId` → `tblPanelRole`
---
### Panel Supporting Tables
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblPanelMemberShipCredentialType | Credential types per membership | PanelMemberShipId, CredentialTypeId |
| tblPanelMemberShipCoordinatorCredentialType | Coordinator credential types | PanelMembershipId, CredentialTypeId |
| tblPanelMemberShipMaterialCredentialType | Material credential types | PanelMemberShipId, CredentialTypeId |
| tblPanelRole | Role definitions | PanelRoleCategoryId |
| tblPanelRoleCategory | Role categories | |
| tblPanelNote | Panel notes | PanelId, NoteId |
---
### Job & Examiner System
### tblJob
Examination marking jobs.
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| JobId | int | | NO | ✅ |
| LanguageId | int | | YES | |
| SentUserID | int | | YES | |
| ReceivedUserId | int | | YES | |
| SentToPayrollUserId | int | | YES | |
| SentDate | datetime | | YES | |
| DueDate | datetime | | YES | |
| ReceivedDate | datetime | | YES | |
| IsComplete | bit | | NO | |
| AcceptedDate | datetime | | YES | |
**FKs:** `LanguageId` → `tblLanguage`, `SentUserID`/`ReceivedUserId`/`SentToPayrollUserId` → `tblUser`
---
### tblJobExaminer
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| JobExaminerID | int | NO | ✅ |
| JobId | int | NO | |
| PanelMembershipId | int | NO | |
| ProductSpecificationId | int | YES | |
| ProductSpecificationChangedUserId | int | YES | |
| ValidatedUserId | int | YES | |
| ValidatedDate | datetime | YES | |
| PaidReviewer | bit | NO | |
**FKs:** `JobId` → `tblJob`, `PanelMembershipId` → `tblPanelMembership`, `ProductSpecificationId` → `tblProductSpecification`
---
### Examiner Supporting Tables
| Table | Key FKs |
|-------|---------|
| tblExaminerMarking | JobExaminerID → tblJobExaminer |
| tblExaminerTestComponentResult | ExaminerMarkingID → tblExaminerMarking, TypeID → tblTestComponentType |
| tblJobExaminerPayroll | JobExaminerId → tblJobExaminer, PayrollId → tblPayroll |
| tblJobExaminerRubricTestComponentResult | JobExaminerId → tblJobExaminer, RubricTestComponentResultId |
| tblExaminerStatusType | (lookup) |
| tblExaminerUnavailable | Examiner availability tracking |
| tblMarkingResultType | (lookup) |
---
## 7. Professional Development
### tblProfessionalDevelopmentActivity
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| ProfessionalDevelopmentActivityId | int | | NO | ✅ |
| PersonId | int | | NO | |
| ProfessionalDevelopmentCategoryId | int | | NO | |
| ProfessionalDevelopmentRequirementId | int | | YES | |
| Description | nvarchar | 500 | NO | |
| Points | decimal | | NO | |
| StartDate | date | | NO | |
| EndDate | date | | YES | |
| Provider | nvarchar | 200 | YES | |
**FKs:** `PersonId` → `tblPerson`, `ProfessionalDevelopmentCategoryId` → `tblProfessionalDevelopmentCategory`, `ProfessionalDevelopmentRequirementId` → `tblProfessionalDevelopmentRequirement`
---
### PD Structure Tables
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblProfessionalDevelopmentCategory | PD categories | ProfessionalDevelopmentCategoryGroupId |
| tblProfessionalDevelopmentCategoryGroup | Category groups | |
| tblProfessionalDevelopmentCategoryRequirement | Category requirements | ProfessionalDevelopmentCategoryId, ProfessionalDevelopmentRequirementId |
| tblProfessionalDevelopmentRequirement | PD requirements | |
| tblProfessionalDevelopmentSection | Sections | |
| tblProfessionalDevelopmentSectionCategory | Section-category links | ProfessionalDevelopmentSectionId, ProfessionalDevelopmentCategoryId, PdPointsLimitTypeId |
| tblProfessionalDevelopmentSubSection | Sub-sections | ProfessionalDevelopmentSectionId |
| tblProfessionalDevelopmentActivityAttachment | Activity file attachments | ProfessionalDevelopmentActivityId, StoredFileId |
| tblProfessionalDevelopmentCredentialApplication | Links PD to applications | CredentialApplicationId, ProfessionalDevelopmentActivityId |
---
## 8. Email & Communication
### tblEmailMessage
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| EmailMessageId | int | | NO | ✅ |
| EmailTemplateId | int | | YES | |
| RecipientEntityId | int | | YES | |
| RecipientUserId | int | | YES | |
| EmailSendStatusTypeId | int | | NO | |
| ToAddress | nvarchar | 200 | NO | |
| FromAddress | nvarchar | 200 | NO | |
| Subject | nvarchar | 500 | NO | |
| Body | nvarchar | max | NO | |
| SentDate | datetime | | YES | |
**FKs:** `EmailTemplateId` → `tblEmailTemplate`, `RecipientEntityId` → `tblEntity`, `RecipientUserId` → `tblUser`, `EmailSendStatusTypeId` → `tblEmailSendStatusType`
---
### Email Supporting Tables
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblEmailTemplate | Reusable templates | |
| tblEmailTemplateDetailType | Template detail types | |
| tblEmailSendStatusType | Send statuses | |
| tblEmailMessageAttachment | Message attachments | EmailMessageId, StoredFileId |
| tblEmailBatch | Batch email operations | UserId, EmailTemplateId |
| tblEmailBatchRecipient | Batch recipients | EmailBatchId, EmailId, PersonId |
| tblEmailChange | Email change requests | UserId → aspnet_Users |
---
### System Action & Email Templates
| Table | Purpose |
|-------|---------|
| tblSystemActionType | Action type definitions |
| tblSystemActionEventType | Event types that trigger actions |
| tblSystemActionEmailTemplate | Links actions to email templates |
| tblSystemActionEmailTemplateDetail | Template detail entries |
---
### Notifications
| Table | Key Columns |
|-------|-------------|
| tblNotification | NotificationId (PK), NotificationTypeId, FromUserId, ToUserId |
| tblNotificationType | NotificationTypeId (PK), Name |
---
## 9. Orders, Payments & Finance
### tblOrder
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| OrderID | int | | NO | ✅ |
| PersonId | int | | YES | |
| OrderDate | datetime | | YES | |
| OrderedUser | int | | YES | |
| InvoiceNumber | varchar | 50 | YES | |
| GSTAmount | money | | YES | |
| TotalAmount | money | | YES | |
| ReferenceNumber | nvarchar | 255 | YES | |
| Verified | bit | | YES | |
---
### tblOrderItem
| Column | Type | Nullable | PK |
|--------|------|----------|-----|
| OrderItemID | int | NO | ✅ |
| OrderID | int | NO | |
| ProductSpecificationId | int | YES | |
| Quantity | int | YES | |
| UnitPrice | money | YES | |
| TotalPrice | money | YES | |
**FKs:** `OrderID` → `tblOrder`
---
### Financial Tables
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblProductSpecification | Product/fee definitions | GLCodeId, JobTypeId, ProductCategoryId |
| tblGLCode | General ledger codes | |
| tblFeeType | Fee type definitions | |
| tblPayroll | Payroll batches | PayrollStatusId → tluPayrollStatus, ModifiedUserId → tblUser |
| tblExternalAccountingOperation | External accounting ops | StatusId, TypeId, RequestedByUserId |
| tblCoupon / tblCouponAssignment / tblCouponRecipient / tblCouponUsage | Coupon/discount system | CouponId |
| tblRefundMethodType | Refund methods | |
---
## 10. File Storage
### tblStoredFile
| Column | Type | Length | Nullable | PK |
|--------|------|--------|----------|-----|
| StoredFileId | int | | NO | ✅ |
| DocumentTypeId | int | | NO | |
| StoredFileStatusTypeId | int | | NO | |
| FileName | nvarchar | 500 | NO | |
| ContentType | nvarchar | 100 | YES | |
| FileSize | int | | YES | |
| StoragePath | nvarchar | 1000 | YES | |
| UploadedByUserId | int | | YES | |
| UploadedByPersonId | int | | YES | |
| UploadedDate | datetime | | NO | |
| Deleted | bit | | NO | |
**FKs:** `DocumentTypeId` → `tluDocumentType`, `StoredFileStatusTypeId` → `tblStoredFileStatusType`, `UploadedByUserId` → `tblUser`, `UploadedByPersonId` → `tblPerson`
---
### File-related Tables
| Table | Purpose |
|-------|---------|
| tblStoredFileStatusType | File status types |
| tblStoredFileDeletePolicy | Deletion policies |
| tblStoredFileDeletePolicyDocumentType | Document types per delete policy |
| tblFile | Legacy file storage |
---
## 11. Security & Access Control
### tblUser
| Column | Type | Length | Nullable | PK | Default |
|--------|------|--------|----------|-----|---------|
| UserId | int | | NO | ✅ | |
| UserName | nvarchar | 50 | NO | | |
| FullName | nvarchar | 100 | NO | | |
| OfficeId | int | | NO | | |
| Note | nvarchar | max | YES | | |
| Active | bit | | NO | | |
| Email | nvarchar | 200 | NO | | |
| SystemUser | bit | | NO | | (0) |
| NonWindowsUser | bit | | NO | | (0) |
| Password | nvarchar | 128 | YES | | |
| LastPasswordChangeDate | datetime | | YES | | |
| FailedPasswordAttemptCount | int | | NO | | (0) |
| IsLockedOut | bit | | NO | | (0) |
| LastLockoutDate | datetime | | YES | | |
**FKs:** `OfficeId` → `tblOffice`
> [!IMPORTANT]
> `tblUser` is the most heavily referenced table in the database — nearly every table has a `ModifiedUser` FK pointing to it.
---
### Security Tables
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblUserRole | User-role assignments | UserId → tblUser, SecurityRoleId → tblSecurityRole |
| tblSecurityRole | Role definitions | |
| tblSecurityNoun | Security nouns (resources) | |
| tblSecurityRule | Rules linking roles to nouns | SecurityRoleId, SecurityNounId |
| tblUserSearch | Saved search configurations | UserId → tblUser |
| tblPasswordHistory | Password history | |
| tblMyNaatiUser | MyNaati portal user links | AspUserId → aspnet_Users |
| tblMyNaatiInvalidCookie | Invalid cookie tracking | |
| tblNcmsInvalidCookie | NCMS invalid cookies | |
| tblApiAccess | API access rules | InstitutionId → tblInstitution |
| tblRegistrationRequest | Registration requests | |
---
## 12. System Configuration
| Table | Purpose |
|-------|---------|
| tblSystemParameter | System-wide parameters |
| tblSystemValue | System configuration values |
| tblTableData | Generic data storage |
| tblDataType | Data type definitions |
| tblApplication | Application metadata |
| tblAuditLog | Audit trail (AuditTypeId → tluAuditType) |
| tblPodHistory | POD history tracking |
| tblDownload | Download tracking |
---
## 13. Location & Geography
### tblVenue
| Column | Type | Length | Nullable | PK | Default |
|--------|------|--------|----------|-----|---------|
| VenueId | int | | NO | ✅ | |
| TestLocationId | int | | NO | | |
| Address | nvarchar | 255 | YES | | |
| Capacity | int | | YES | | |
| Name | nvarchar | 100 | NO | | |
| PublicNotes | nvarchar | 1000 | YES | | |
| Inactive | bit | | NO | | (0) |
| Coordinates | varchar | 22 | YES | | |
**FKs:** `TestLocationId` → `tblTestLocation`
---
### Geography Tables
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblTestLocation | Test locations (cities/regions) | CountryId → tblCountry, OfficeId → tblOffice |
| tblOffice | NAATI offices | InstitutionId → tblInstitution, StateId → tluState |
| tblLocation | General locations | CountryId → tblCountry, StateId → tluState |
| tblCountry | Countries | |
| tblSuburb | Suburbs | StateId → tluState |
| tblPostcode | Postcodes | SuburbId → tblSuburb |
---
## 14. Role Players
Role players participate in oral/dialogue test sessions.
| Table | Purpose | Key FKs |
|-------|---------|---------|
| tblRolePlayer | Role player records | PersonId → tblPerson |
| tblRolePlayerTestLocation | Available locations | RolePlayerId, TestLocationId |
| tblRolePlayerRoleType | Role type definitions | |
| tblRolePlayerStatusType | Status type definitions | |
| tblTestSessionRolePlayer | Session assignments | RolePlayerId, TestSessionId, RolePlayerStatusTypeId, StatusChangeUserId |
| tblTestSessionRolePlayerDetail | Assignment details | TestSessionRolePlayerId, TestComponentId, SkillId, LanguageId, RolePlayerRoleTypeId |
---
## 15. Lookup Tables (tlu*)
Small reference/lookup tables with standardized structure:
| Table | PK | Key Columns |
|-------|----|-------------|
| tluAccessDisabledType | AccessDisabledTypeId | Name |
| tluAuditType | AuditTypeId | Name |
| tluDocumentType | DocumentTypeId | Name, DisplayName, ExaminerToolsDownload/Upload, DocumentTypeCategoryId |
| tluEFTMachine | EFTMachineId | TerminalNo, OfficeId, Visible |
| tluExternalAccountingOperationStatus | ExternalAccountingOperationStatusId | Name, DisplayName, Description |
| tluExternalAccountingOperationType | ExternalAccountingOperationTypeId | Name, DisplayName |
| tluJobType | JobTypeId | Name |
| tluPanelType | PanelTypeId | Name, Description, AllowCredentialTypeLink |
| tluPayrollStatus | PayrollStatusId | Name, DisplayName, Description |
| tluPermission | PermissionId | Name, Description, Rank |
| tluProductCategory | ProductCategoryId | Name, Code, ProductTypeId → tluProductType, DisplayName |
| tluProductType | ProductTypeId | Name, DisplayName |
| tluRegion | RegionId | Name, StateId → tluState, CountryId → tblCountry |
| tluResultType | ResultTypeId | Result |
| tluState | StateId | State (abbreviation), Name |
| tluTitle | TitleId | Title, StandardTitle |
---
## 16. HangFire Job Queue
Background job processing tables (present in both `HangFire` and `MyNaatiHangFire` schemas):
| Table | PK | Purpose |
|-------|----|---------|
| AggregatedCounter | Id | Aggregated counter values |
| Counter | Id | Raw counter values |
| Hash | Id | Key-value hash storage |
| Job | Id | Job definitions — FKs from JobParameter and State |
| JobParameter | Id | Job parameters (JobId → Job) |
| JobQueue | Id | Job queue entries |
| List | Id | List storage |
| Schema | Version | Schema version tracking |
| Server | Id | Server registration |
| Set | Id | Set storage |
| State | Id | Job state history (JobId → Job) |
---
## 17. Miscellaneous
| Table | Purpose |
|-------|---------|
| Users | Secondary user table (id: char(36) PK, email, password, name, customerNumber) |
| VersionInfo | Database migration version tracking |
| tblLegacyAccreditation | Legacy accreditation data |
| tblRecertification | Recertification records — CertificationPeriodId, CredentialApplicationId |
| tblWorkPractice | Work practice records — CredentialId |
| tblWorkPracticeAttachment | Work practice file attachments |
| tblWorkPracticeCredentialRequest | Links work practice to credential requests |
| tblCandidateBrief | Candidate brief documents — TestSittingId, TestMaterialAttachmentId |
| tblSubmitTestDraft | Draft test submissions |
| tblSubmitTestDraftAttachment | Draft attachments — SubmitTestDraftID |
| tblSubmitTestDraftComponent | Draft components — SubmitTestDraftID |
---
## 18. Database Views
| View | Purpose |
|------|---------|
| vwDistinctInstitutionName | Distinct institution names |
| vwDistinctPersonName | Distinct person names |
| vwIssuedCredentialCredentialRequest | Links issued credentials to requests |
| vwJobExaminerPayrollStatus | Examiner payroll status |
| vwMarkingsForPayroll | Markings ready for payroll processing |
| vwMaxPersonNameEffectiveDate | Latest person name effective date |
| vwPersonDistinct | Distinct person records with names |
| vwRolePlayerLastTestSession | Last test session per role player |
| vwSkillDisplayName | Formatted skill display names |
| vwTestMaterialCreationPayments | Material creation payment tracking |
| vwTestMaterialLastUsed | Last used date per test material |
| vwTestMaterialRequestRoundLatest | Latest round per material request |
| vwTestStatus | Comprehensive test status with flags |
*Plus 5 legacy ASP.NET views: `vw_aspnet_Applications`, `vw_aspnet_MembershipUsers`, `vw_aspnet_Roles`, `vw_aspnet_Users`, `vw_aspnet_UsersInRoles`*
---
## Entity Relationship Summary
```mermaid
graph TD
    E[tblEntity] --> P[tblPerson]
    E --> I[tblInstitution]
    E --> ADDR[tblAddress]
    E --> PHONE[tblPhone]
    E --> EMAIL[tblEmail]
    
    P --> PN[tblPersonName]
    P --> CP[tblCertificationPeriod]
    P --> PM[tblPanelMembership]
    P --> PD[tblProfessionalDevelopmentActivity]
    
    P --> CA[tblCredentialApplication]
    CA --> CR[tblCredentialRequest]
    CR --> CRED[tblCredential]
    
    CR --> TS[tblTestSitting]
    TS --> TSESS[tblTestSession]
    TSESS --> V[tblVenue]
    V --> TL[tblTestLocation]
    
    TS --> TR[tblTestResult]
    TR --> J[tblJob]
    J --> JE[tblJobExaminer]
    JE --> PM
    
    CRED --> CT[tblCredentialType]
    CR --> SK[tblSkill]
    SK --> L[tblLanguage]
    
    TSESS --> TSPEC[tblTestSpecification]
    TSPEC --> TCT[tblTestComponentType]
    TCT --> TC[tblTestComponent]
    
    CA --> CAT[tblCredentialApplicationType]
    CA --> OFF[tblOffice]
    OFF --> I
    
    PM --> PAN[tblPanel]
    PAN --> L
TIP

The most central tables are: tblEntity (root for all persons/institutions), tblPerson (individuals), tblCredentialApplication (application workflow hub), tblCredentialRequest (links applications to credentials), tblTestSitting (test participation), and tblUser (administrative users — referenced by nearly every table via ModifiedUser FK).