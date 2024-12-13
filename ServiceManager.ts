import { ServiceScope } from "@microsoft/sp-core-library";
import { SPService } from "./SPService";
/**
 * The ServiceManager class creates a singleton of our custom services so that
 * we can easily access these services anywhere within our application without
 * needing to recreate the services and pass the serviceScope around the app.
 */
class ServiceManager {
  private static _spService: SPService;
  // Initialize the service with serviceScope
  public static initialize(serviceScope: ServiceScope): void {
    this._spService = serviceScope.consume(SPService.serviceKey);
  }
  // Getter to access the SPListService
  public static get spService(): SPService {
    if (!this._spService) {
      throw new Error("ServiceManager not initialized with serviceScope");
    }
    return this._spService;
  }
}
export default ServiceManager;
