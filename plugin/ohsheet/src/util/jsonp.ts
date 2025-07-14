interface JsonpOptions {
  url: string;
  data?: any;
  timeout?: number;
}

interface JsonpResponse<T = any> {
  success: boolean;
  data?: T;
  error?: string;
}

export function jsonpRequest<T = any>(options: JsonpOptions): Promise<JsonpResponse<T>> {
  return new Promise((resolve) => {
    const { url, data, timeout = 10000 } = options;
    
    // Create a unique callback name
    const callbackName = 'jsonpCallback_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    
    // Create the script element
    const script = document.createElement('script');
    
    // Set up the callback function
    (window as any)[callbackName] = (result: T) => {
      cleanup();
      resolve({ success: true, data: result });
    };
    
    // Build the URL with the callback parameter
    const urlObj = new URL(url);
    urlObj.searchParams.set('callback', callbackName);
    
    if (data) {
      urlObj.searchParams.set('data', JSON.stringify(data));
    }
    
    script.src = urlObj.toString();
    
    // Add error handling
    script.onerror = () => {
      cleanup();
      resolve({ success: false, error: 'Failed to load script' });
    };
    
    // Add timeout handling
    const timeoutId = setTimeout(() => {
      cleanup();
      resolve({ success: false, error: 'Request timeout' });
    }, timeout);
    
    // Cleanup function
    const cleanup = () => {
      clearTimeout(timeoutId);
      delete (window as any)[callbackName];
      if (script.parentNode) {
        script.parentNode.removeChild(script);
      }
    };
    
    // Add the script to the document
    document.head.appendChild(script);
  });
}

// Convenience function for POST-like requests
export function jsonpPost<T = any>(url: string, data: any, timeout?: number): Promise<JsonpResponse<T>> {
  return jsonpRequest<T>({ url, data, timeout });
}

// Convenience function for GET-like requests
export function jsonpGet<T = any>(url: string, timeout?: number): Promise<JsonpResponse<T>> {
  return jsonpRequest<T>({ url, timeout });
} 