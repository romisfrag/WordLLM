function setInLocalStorage(key: string, value: string) {
    const myPartitionKey = Office.context.partitionKey;
  
    // Check if local storage is partitioned. 
    // If so, use the partition to ensure the data is only accessible by your add-in.
    if (myPartitionKey) {
      localStorage.setItem(myPartitionKey + key, value);
    } else {
      localStorage.setItem(key, value);
    }
  }
  
  function getFromLocalStorage(key: string) {
    const myPartitionKey = Office.context.partitionKey;
  
    // Check if local storage is partitioned.
    if (myPartitionKey) {
      return localStorage.getItem(myPartitionKey + key);
    } else {
      return localStorage.getItem(key);
    }
  }

  export { setInLocalStorage, getFromLocalStorage };