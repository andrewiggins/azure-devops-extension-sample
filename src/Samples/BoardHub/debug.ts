import {
  CommonServiceIds,
  IHostNavigationService,
  ILocationService,
  IProjectPageService,
} from "azure-devops-extension-api";
import * as SDK from "azure-devops-extension-sdk";

export async function getAllSDKDebugInfo() {
  return {
    sdkProps: await getAllStaticHostProperties(),
    navService: await getHostNavigationProperties(),
    projectPage: await getProjectPageProperties(),
  };
}

export async function getAllStaticHostProperties() {
  await SDK.ready();

  return {
    getConfiguration: SDK.getConfiguration(),
    getContributionId: SDK.getContributionId(),
    getExtensionContext: SDK.getExtensionContext(),
    getHost: SDK.getHost(),
    getUser: SDK.getUser(),
    sdkVersion: SDK.sdkVersion,
  };
}

export async function getHostNavigationProperties() {
  await SDK.ready();

  const navService = await SDK.getService<IHostNavigationService>(
    CommonServiceIds.HostNavigationService
  );

  return {
    getHash: await navService.getHash(),
    getPageNavigationElements: await navService.getPageNavigationElements(),
    getPageRoute: await navService.getPageRoute(),
    getQueryParams: await navService.getQueryParams(),
  };
}

// TODO: Requires inputs I don't understand yet
// export async function getLocationServiceProperties() {
//	await SDK.ready();
//
// 	const locationService = await SDK.getService<ILocationService>(
// 		CommonServiceIds.LocationService
// 	);
//
// 	return {};
// }

export async function getProjectPageProperties() {
  await SDK.ready();

  const projectPage = await SDK.getService<IProjectPageService>(
    CommonServiceIds.ProjectPageService
  );

  return { project: await projectPage.getProject() };
}
