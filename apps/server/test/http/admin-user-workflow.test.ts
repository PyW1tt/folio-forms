import { afterEach, test } from "bun:test";

import type { PrefillMockServer } from "prefill-mock/mock";

import { createApp } from "../../src/app";
import {
  runAccountAuthority,
  runResponseErasure,
} from "../scenarios/admin-workflow/account-erasure";
import {
  runPasswordAndThrottle,
  runManagedAccounts,
} from "../scenarios/admin-workflow/account-lifecycle";
import { runCallbackSecurity } from "../scenarios/admin-workflow/callback-security";
import {
  runCorrection,
  runCorrectionFailureAndRevisions,
} from "../scenarios/admin-workflow/correction";
import { runDraftLifecycle } from "../scenarios/admin-workflow/draft";
import { runDraftResultsAndExports } from "../scenarios/admin-workflow/draft-results";
import {
  runFieldConfiguration,
  runEditorAccess,
} from "../scenarios/admin-workflow/editor-access";
import {
  runErasureFailures,
  runImmutableAudits,
} from "../scenarios/admin-workflow/erasure-audit";
import { runFormLifecycle } from "../scenarios/admin-workflow/form-lifecycle";
import { runOfficePictureResponse } from "../scenarios/admin-workflow/picture-response";
import { runPrefillContinuation } from "../scenarios/admin-workflow/prefill-continuation";
import { runPrefillEntry } from "../scenarios/admin-workflow/prefill-entry";
import {
  runDiscardAndRestart,
  runRecoverableState,
} from "../scenarios/admin-workflow/recovery";
import {
  runBootstrapAndCreation,
  runTemplateUploads,
} from "../scenarios/admin-workflow/setup";
import {
  runSubmitFailuresAndCompletion,
  runSubmissionExports,
} from "../scenarios/admin-workflow/submission";
import {
  runPrimaryPublication,
  runValidTemplateContracts,
  runInvalidTemplateContracts,
} from "../scenarios/admin-workflow/template-contract";

// oxlint-disable no-await-in-loop complexity -- The end-to-end journey deliberately keeps sequential transitions in one test.
const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) {
  throw new Error(
    "The HTTP application test requires DATABASE_URL for an isolated PostgreSQL database"
  );
}
const convertedDocumentKeys: string[] = [];
const app = createApp({
  legacySso: null,
  onlyOffice: {
    convertDocxToPdf: (documentKey) => {
      convertedDocumentKeys.push(documentKey);
      return Promise.resolve(new TextEncoder().encode("%PDF-test"));
    },
    forceSave: () => Promise.resolve(false),
  },
  prefillReturnUrl: "https://source.example.test/forms/return",
  requestIp: (request) => request.headers.get("x-test-ip"),
});

let externalMock: PrefillMockServer | undefined;

afterEach(() => {
  externalMock?.close();
  externalMock = undefined;
});

test("serves authenticated Admin and User workflows through HTTP", async () => {
  const bootstrapAndCreation = await runBootstrapAndCreation({
    app,
  });
  const templateUploads = await runTemplateUploads({
    adminBearer: bootstrapAndCreation.adminBearer,
    adminEmail: bootstrapAndCreation.adminEmail,
    app,
    password: bootstrapAndCreation.password,
  });
  const fieldConfiguration = await runFieldConfiguration({
    adminBearer: bootstrapAndCreation.adminBearer,
    adminId: bootstrapAndCreation.adminId,
    app,
    createdFormRecord: bootstrapAndCreation.createdFormRecord,
    formId: bootstrapAndCreation.formId,
    formPublicId: bootstrapAndCreation.formPublicId,
    templateDocumentKey: bootstrapAndCreation.templateDocumentKey,
  });
  const editorAccess = await runEditorAccess({
    adminBearer: bootstrapAndCreation.adminBearer,
    adminEditor: fieldConfiguration.adminEditor,
    adminId: bootstrapAndCreation.adminId,
    adminLease: fieldConfiguration.adminLease,
    app,
    competingAdminBearer: templateUploads.competingAdminBearer,
    createdFormRecord: bootstrapAndCreation.createdFormRecord,
    formId: bootstrapAndCreation.formId,
    formPublicId: bootstrapAndCreation.formPublicId,
    publishCapability: fieldConfiguration.publishCapability,
    saveTemplateCapability: fieldConfiguration.saveTemplateCapability,
    templateDocumentKey: bootstrapAndCreation.templateDocumentKey,
  });
  const primaryPublication = await runPrimaryPublication({
    adminBearer: bootstrapAndCreation.adminBearer,
    app,
    formId: bootstrapAndCreation.formId,
    formPublicId: bootstrapAndCreation.formPublicId,
    initialTemplateBytes: editorAccess.initialTemplateBytes,
    saveTemplateCapability: editorAccess.saveTemplateCapability,
    selectedPointer: fieldConfiguration.selectedPointer,
    templateDocumentKey: bootstrapAndCreation.templateDocumentKey,
  });
  await runValidTemplateContracts({
    adminBearer: bootstrapAndCreation.adminBearer,
    app,
  });
  const officePictureResponse = await runOfficePictureResponse({
    adminBearer: bootstrapAndCreation.adminBearer,
    app,
  });
  await runInvalidTemplateContracts({
    activePublishCapability: primaryPublication.activePublishCapability,
    activeTemplateDocumentKey: primaryPublication.activeTemplateDocumentKey,
    adminBearer: bootstrapAndCreation.adminBearer,
    app,
    formPublicId: bootstrapAndCreation.formPublicId,
    refreshedAdminEditor: primaryPublication.refreshedAdminEditor,
  });
  const prefillEntry = await runPrefillEntry({
    app,
    formId: bootstrapAndCreation.formId,
    password: bootstrapAndCreation.password,
    selectedPointer: fieldConfiguration.selectedPointer,
    userEmail: bootstrapAndCreation.userEmail,
  });
  externalMock = prefillEntry.mock;
  await runPrefillContinuation({
    app,
    formId: bootstrapAndCreation.formId,
    formRecord: prefillEntry.formRecord,
    handoffCandidateValues: prefillEntry.handoffCandidateValues,
    handoffCreateBody: prefillEntry.handoffCreateBody,
    missingValueBearer: prefillEntry.missingValueBearer,
    mock: prefillEntry.mock,
    password: bootstrapAndCreation.password,
    prefillHandoffSecret: prefillEntry.prefillHandoffSecret,
    secondFormPublicId: editorAccess.secondFormPublicId,
    userBearer: prefillEntry.userBearer,
    userEmail: bootstrapAndCreation.userEmail,
  });
  const formLifecycle = await runFormLifecycle({
    adminBearer: bootstrapAndCreation.adminBearer,
    adminId: bootstrapAndCreation.adminId,
    app,
    formId: bootstrapAndCreation.formId,
    formPublicId: bootstrapAndCreation.formPublicId,
    formRecord: prefillEntry.formRecord,
    password: bootstrapAndCreation.password,
    publishedBytes: primaryPublication.publishedBytes,
    secretFormDescription: bootstrapAndCreation.secretFormDescription,
    secretFormTitle: bootstrapAndCreation.secretFormTitle,
    user: prefillEntry.user,
    userBearer: prefillEntry.userBearer,
  });
  const draftLifecycle = await runDraftLifecycle({
    adminBearer: bootstrapAndCreation.adminBearer,
    app,
    formId: bootstrapAndCreation.formId,
    formPublicId: bootstrapAndCreation.formPublicId,
    formRecord: prefillEntry.formRecord,
    secretFormTitle: bootstrapAndCreation.secretFormTitle,
    selectedPointer: fieldConfiguration.selectedPointer,
    userBearer: prefillEntry.userBearer,
    userId: prefillEntry.userId,
  });
  await runDraftResultsAndExports({
    adminBearer: bootstrapAndCreation.adminBearer,
    adminId: bootstrapAndCreation.adminId,
    app,
    formId: bootstrapAndCreation.formId,
    formRecord: prefillEntry.formRecord,
    publishedManifestRecord: primaryPublication.publishedManifestRecord,
    responseId: draftLifecycle.responseId,
    savedDraftData: draftLifecycle.savedDraftData,
    savedResponseDocumentBytes: draftLifecycle.savedResponseDocumentBytes,
    savedResponseDocumentKey: draftLifecycle.savedResponseDocumentKey,
    savedResponseObjectKey: draftLifecycle.savedResponseObjectKey,
    userBearer: prefillEntry.userBearer,
    userEmail: bootstrapAndCreation.userEmail,
  });
  const submitFailuresAndCompletion = await runSubmitFailuresAndCompletion({
    adminBearer: bootstrapAndCreation.adminBearer,
    app,
    formId: bootstrapAndCreation.formId,
    formPublicId: bootstrapAndCreation.formPublicId,
    formRecord: prefillEntry.formRecord,
    handoffCandidateValues: prefillEntry.handoffCandidateValues,
    mock: prefillEntry.mock,
    responseDocumentKey: draftLifecycle.responseDocumentKey,
    responseId: draftLifecycle.responseId,
    saveBody: draftLifecycle.saveBody,
    saveDraftCapability: draftLifecycle.saveDraftCapability,
    savedDraftData: draftLifecycle.savedDraftData,
    submitCapability: draftLifecycle.submitCapability,
    userBearer: prefillEntry.userBearer,
    userEmail: bootstrapAndCreation.userEmail,
    userId: prefillEntry.userId,
  });
  const submissionExports = await runSubmissionExports({
    adminBearer: bootstrapAndCreation.adminBearer,
    app,
    completedSubmissionId: submitFailuresAndCompletion.completedSubmissionId,
    formId: bootstrapAndCreation.formId,
    formRecord: prefillEntry.formRecord,
    otherUserEmail: bootstrapAndCreation.otherUserEmail,
    password: bootstrapAndCreation.password,
    publishedManifest: primaryPublication.publishedManifest,
    publishedManifestRecord: primaryPublication.publishedManifestRecord,
    responseId: draftLifecycle.responseId,
    savedDraftData: draftLifecycle.savedDraftData,
    userBearer: prefillEntry.userBearer,
  });
  const correction = await runCorrection({
    adminBearer: bootstrapAndCreation.adminBearer,
    adminId: bootstrapAndCreation.adminId,
    app,
    competingAdminBearer: templateUploads.competingAdminBearer,
    completedSubmissionId: submitFailuresAndCompletion.completedSubmissionId,
    dataBody: submissionExports.dataBody,
    formId: bootstrapAndCreation.formId,
    formRecord: prefillEntry.formRecord,
    mock: prefillEntry.mock,
    otherUserBearer: submissionExports.otherUserBearer,
    postSubmitExternalReference:
      submitFailuresAndCompletion.postSubmitExternalReference,
    responseDocumentKey: submitFailuresAndCompletion.responseDocumentKey,
    responseId: draftLifecycle.responseId,
    savedDraftData: draftLifecycle.savedDraftData,
    userBearer: prefillEntry.userBearer,
  });
  await runCorrectionFailureAndRevisions({
    adminBearer: bootstrapAndCreation.adminBearer,
    adminId: bootstrapAndCreation.adminId,
    app,
    completedSubmissionId: submitFailuresAndCompletion.completedSubmissionId,
    convertedDocumentKeys,
    correction: correction.correction,
    correctionData: correction.correctionData,
    originalSubmissionBeforeCorrection:
      correction.originalSubmissionBeforeCorrection,
    otherUserBearer: submissionExports.otherUserBearer,
    responseId: draftLifecycle.responseId,
    savedDraftData: draftLifecycle.savedDraftData,
    submissionDocument: submissionExports.submissionDocument,
    userBearer: prefillEntry.userBearer,
    userId: prefillEntry.userId,
  });
  await runDiscardAndRestart({
    adminBearer: bootstrapAndCreation.adminBearer,
    adminId: bootstrapAndCreation.adminId,
    app,
    archivedNoResponseBearer: formLifecycle.archivedNoResponseBearer,
    archivedNoResponseEmail: formLifecycle.archivedNoResponseEmail,
    formId: bootstrapAndCreation.formId,
    formPublicId: bootstrapAndCreation.formPublicId,
    handoffCandidateValues: prefillEntry.handoffCandidateValues,
    prefillHandoffSecret: prefillEntry.prefillHandoffSecret,
    publishedContractBefore: formLifecycle.publishedContractBefore,
    userBearer: prefillEntry.userBearer,
  });
  await runRecoverableState({
    adminBearer: bootstrapAndCreation.adminBearer,
    adminId: bootstrapAndCreation.adminId,
    app,
    formId: bootstrapAndCreation.formId,
    formPublicId: bootstrapAndCreation.formPublicId,
    otherUser: submissionExports.otherUser,
    sourceDocument: submissionExports.sourceDocument,
    templateDraft: submissionExports.templateDraft,
  });
  await runCallbackSecurity({
    adminBearer: bootstrapAndCreation.adminBearer,
    adminId: bootstrapAndCreation.adminId,
    formId: bootstrapAndCreation.formId,
    formPublicId: bootstrapAndCreation.formPublicId,
    publishCapability: fieldConfiguration.publishCapability,
    publishedManifestRecord: primaryPublication.publishedManifestRecord,
    sourceDocument: submissionExports.sourceDocument,
  });
  await runPasswordAndThrottle({
    app,
    formRecord: prefillEntry.formRecord,
  });
  const managedAccounts = await runManagedAccounts({
    adminBearer: bootstrapAndCreation.adminBearer,
    app,
  });
  const accountAuthority = await runAccountAuthority({
    adminBearer: bootstrapAndCreation.adminBearer,
    app,
  });
  const responseErasure = await runResponseErasure({
    app,
    erasureAdminBearer: accountAuthority.erasureAdminBearer,
    erasureAdminId: accountAuthority.erasureAdminId,
    formRecord: prefillEntry.formRecord,
    mock: prefillEntry.mock,
    pictureCorrectionEditor: officePictureResponse.pictureCorrectionEditor,
    pictureResponseId: officePictureResponse.pictureResponseId,
    postSubmitExternalReference:
      submitFailuresAndCompletion.postSubmitExternalReference,
    postSubmitHandoffRecord:
      submitFailuresAndCompletion.postSubmitHandoffRecord,
    responseDocumentKey: submitFailuresAndCompletion.responseDocumentKey,
    responseId: draftLifecycle.responseId,
    saveDraftCapability: submitFailuresAndCompletion.saveDraftCapability,
    savedDraftData: draftLifecycle.savedDraftData,
    userBearer: prefillEntry.userBearer,
    userEmail: bootstrapAndCreation.userEmail,
    userId: prefillEntry.userId,
    userLease: submitFailuresAndCompletion.userLease,
  });
  await runErasureFailures({
    app,
    erasureAdminBearer: accountAuthority.erasureAdminBearer,
    formRecord: prefillEntry.formRecord,
    handoffCandidateValues: prefillEntry.handoffCandidateValues,
    otherUser: submissionExports.otherUser,
    otherUserBearer: submissionExports.otherUserBearer,
    otherUserEmail: bootstrapAndCreation.otherUserEmail,
    password: bootstrapAndCreation.password,
    prefillHandoffSecret: prefillEntry.prefillHandoffSecret,
  });
  await runImmutableAudits({
    adminBearer: bootstrapAndCreation.adminBearer,
    app,
    authorityAdminA: accountAuthority.authorityAdminA,
    authorityAdminB: accountAuthority.authorityAdminB,
    competingAdmin: templateUploads.competingAdmin,
    competingAdminBearer: templateUploads.competingAdminBearer,
    createdFormRecord: bootstrapAndCreation.createdFormRecord,
    erasureAdminBearer: accountAuthority.erasureAdminBearer,
    erasureAdminId: accountAuthority.erasureAdminId,
    formPublicId: bootstrapAndCreation.formPublicId,
    invalidTargetSecret: managedAccounts.invalidTargetSecret,
    mainTombstone: responseErasure.mainTombstone,
    managedId: managedAccounts.managedId,
    managedPassword: managedAccounts.managedPassword,
    managedTemporaryPassword: managedAccounts.managedTemporaryPassword,
    managedToken: managedAccounts.managedToken,
    otherUserEmail: bootstrapAndCreation.otherUserEmail,
    password: bootstrapAndCreation.password,
    pictureResponseId: officePictureResponse.pictureResponseId,
    postSubmitExternalReference:
      submitFailuresAndCompletion.postSubmitExternalReference,
    publishCapability: fieldConfiguration.publishCapability,
    resetPassword: managedAccounts.resetPassword,
    resetTemporaryPassword: managedAccounts.resetTemporaryPassword,
    saveDraftCapability: submitFailuresAndCompletion.saveDraftCapability,
    saveTemplateCapability: editorAccess.saveTemplateCapability,
    savedDraftData: draftLifecycle.savedDraftData,
    secretFormDescription: bootstrapAndCreation.secretFormDescription,
    secretFormTitle: bootstrapAndCreation.secretFormTitle,
    submitCapability: submitFailuresAndCompletion.submitCapability,
    templateDocumentKey: bootstrapAndCreation.templateDocumentKey,
    uploadPublicId: templateUploads.uploadPublicId,
    userBearer: prefillEntry.userBearer,
  });
});
