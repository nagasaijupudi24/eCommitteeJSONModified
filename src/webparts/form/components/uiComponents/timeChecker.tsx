/* eslint-disable @typescript-eslint/no-non-null-assertion */
export  class TimeCheckerTab {
    private static readonly UNIQUE_KEY_PREFIX = "tabTimeKey_"; // Prefix for the unique key
    private static uniqueKey: string | null = null; // Unique key for this tab
    private static readonly TIME_THRESHOLD = 20000; // 20 seconds in milliseconds
  
    /**
     * Initialize time checking logic for page load and tab visibility change.
     */
    public static initPageLoadTimeCheck = (): void => {
      console.log("Initializing time checker for page load...");
  
      // Generate and store the unique key when the page loads
      this.storeUniqueKeyOnPageLoad();
  
      // Add event listener for tab visibility change
      document.addEventListener("visibilitychange", this.handleVisibilityChange);
      console.log("Event listener for 'visibilitychange' added.");
    };
  
    /**
     * Handle tab visibility change events.
     */
    private static handleVisibilityChange = (): void => {
      console.log("Visibility change detected. Current visibility state:", document.visibilityState);
  
      if (document.visibilityState === "visible") {
        // Compare the current time with the stored timestamp
        const storedTime = localStorage.getItem(this.uniqueKey!); // Key will be non-null after storage
        const currentTime = new Date().getTime();
  
        if (storedTime) {
          const timeDifference = currentTime - parseInt(storedTime, 10);
          console.log("Time difference since key storage:", timeDifference, "ms");
  
          if (timeDifference > this.TIME_THRESHOLD) { // Check if more than 20 seconds
            console.log("User returned to tab after more than 20 seconds. Reloading the page...");
            // Clean up this tab's key from localStorage before reloading
            localStorage.removeItem(this.uniqueKey!);
            window.location.reload();
          } else {
            console.log("Time difference is less than 20 seconds. No action taken.");
          }
        } else {
          console.log("No timestamp found for this tab in localStorage.");
        }
      }
    };
  
    /**
     * Generate and store a unique key in localStorage on page load.
     */
    private static storeUniqueKeyOnPageLoad = (): void => {
      if (!this.uniqueKey) {
        const currentTime = new Date().getTime();
        this.uniqueKey = `${this.UNIQUE_KEY_PREFIX}${Date.now()}`; // Generate unique key
        localStorage.setItem(this.uniqueKey, currentTime.toString());
        console.log(`Unique key "${this.uniqueKey}" stored with timestamp:`, currentTime);
      } else {
        console.log("Unique key already exists. No need to store again.");
      }
    };
  }
  
  
  
  






export default class TimeChecker {
    private lastCheckedTime: number;
  
    constructor() {
      this.lastCheckedTime = new Date().getTime();

      console.log(this.lastCheckedTime)
    }
  
    public startTimeCheck(interval: number = 60000): void {
      // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
      const checkAndSchedule = () => {
        this.checkTimeDifferenceAndRefresh();
        setTimeout(checkAndSchedule, interval);
      };
  
      checkAndSchedule(); // Initial call
    }
  
    private checkTimeDifferenceAndRefresh(): void {
      const currentTime = new Date().getTime();
      const timeDifference = currentTime - this.lastCheckedTime;
      console.log(timeDifference)
  
      if (timeDifference > 60000) {
        console.log("Time difference exceeded 1 minute. Refreshing the page...");
        // Optimize to refresh only the relevant part of the page if needed
        window.location.reload();
      }
  
      this.lastCheckedTime = currentTime;
    }
  }
  