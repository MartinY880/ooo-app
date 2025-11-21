import React, { useState, useEffect } from 'react';
import { useMsal } from '@azure/msal-react';
import { graphConfig } from '../authConfig';
import { useDebounce } from '../hooks/useDebounce';

/**
 * Interface for Microsoft Graph User object
 * Contains the essential properties we need from the Graph API
 */
interface GraphUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
}

interface UserSearchProps {
  onUserSelect: (user: GraphUser | null) => void;
  selectedUser: GraphUser | null;
}

/**
 * UserSearch Component
 * 
 * Provides an autocomplete search field that queries Microsoft Graph API
 * to find users in the organization. Uses debouncing to prevent excessive
 * API calls while the user is typing.
 */
export const UserSearch: React.FC<UserSearchProps> = ({ onUserSelect, selectedUser }) => {
  const { instance, accounts } = useMsal();
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [results, setResults] = useState<GraphUser[]>([]);
  const [isSearching, setIsSearching] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [showDropdown, setShowDropdown] = useState<boolean>(false);

  // Debounce the search term to avoid excessive API calls
  const debouncedSearchTerm = useDebounce(searchTerm, 400);

  /**
   * Search for users in Microsoft Graph
   * Uses the $search query parameter which requires ConsistencyLevel: eventual header
   */
  useEffect(() => {
    const searchUsers = async () => {
      // Don't search if a user is already selected
      if (selectedUser) {
        return;
      }

      // Only search if we have a valid search term (minimum 2 characters)
      if (debouncedSearchTerm.length < 2) {
        setResults([]);
        setShowDropdown(false);
        return;
      }

      setIsSearching(true);
      setError('');

      try {
        // Acquire an access token silently for Microsoft Graph API
        const response = await instance.acquireTokenSilent({
          scopes: graphConfig.scopes,
          account: accounts[0],
        });

        // Make the Graph API request with the access token
        // Important: The $search parameter requires the ConsistencyLevel: eventual header
        const graphResponse = await fetch(
          `${graphConfig.graphUsersEndpoint}?$search="displayName:${debouncedSearchTerm}" OR "mail:${debouncedSearchTerm}"&$filter=userType eq 'Member' and accountEnabled eq true&$top=10&$select=id,displayName,mail,userPrincipalName,jobTitle,department`,
          {
            headers: {
              Authorization: `Bearer ${response.accessToken}`,
              'ConsistencyLevel': 'eventual', // Required for $search queries
            },
          }
        );

        if (!graphResponse.ok) {
          throw new Error(`Graph API error: ${graphResponse.statusText}`);
        }

        const data = await graphResponse.json();
        setResults(data.value || []);
        setShowDropdown(true);
      } catch (err) {
        console.error('Error searching users:', err);
        setError('Failed to search users. Please try again.');
        setResults([]);
      } finally {
        setIsSearching(false);
      }
    };

    searchUsers();
  }, [debouncedSearchTerm, instance, accounts, selectedUser]);

  /**
   * Handle user selection from the dropdown
   */
  const handleUserClick = (user: GraphUser) => {
    setSearchTerm(user.displayName);
    setResults([]);
    setShowDropdown(false);
    onUserSelect(user);
  };

  /**
   * Handle clearing the selected user
   */
  const handleClear = () => {
    setSearchTerm('');
    setResults([]);
    setShowDropdown(false);
    onUserSelect(null);
  };

  return (
    <div className="relative">
      <label htmlFor="userSearch" className="block text-sm font-medium text-gray-700 mb-1">
        Forward To
      </label>
      
      <div className="relative">
        <input
          id="userSearch"
          type="text"
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          onFocus={() => results.length > 0 && setShowDropdown(true)}
          placeholder="Search for a user by name or email..."
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
        />
        
        {/* Loading indicator */}
        {isSearching && (
          <div className="absolute right-10 top-2.5">
            <div className="animate-spin h-5 w-5 border-2 border-blue-500 border-t-transparent rounded-full"></div>
          </div>
        )}
        
        {/* Clear button */}
        {searchTerm && (
          <button
            type="button"
            onClick={handleClear}
            className="absolute right-3 top-2.5 text-gray-400 hover:text-gray-600"
          >
            ✕
          </button>
        )}
      </div>

      {/* Error message */}
      {error && (
        <p className="mt-1 text-sm text-red-600">{error}</p>
      )}

      {/* Selected user display */}
      {selectedUser && !showDropdown && (
        <div className="mt-2 p-3 bg-blue-50 border border-blue-200 rounded-lg">
          <div className="flex justify-between items-start">
            <div>
              <p className="font-medium text-gray-900">{selectedUser.displayName}</p>
              <p className="text-sm text-gray-600">{selectedUser.mail || selectedUser.userPrincipalName}</p>
              {selectedUser.jobTitle && (
                <p className="text-xs text-gray-500">{selectedUser.jobTitle}</p>
              )}
            </div>
            <button
              type="button"
              onClick={handleClear}
              className="text-gray-400 hover:text-gray-600"
            >
              ✕
            </button>
          </div>
        </div>
      )}

      {/* Dropdown results */}
      {showDropdown && results.length > 0 && (
        <div className="absolute z-10 w-full mt-1 bg-white border border-gray-300 rounded-lg shadow-lg max-h-60 overflow-auto">
          {results.map((user) => (
            <button
              key={user.id}
              type="button"
              onClick={() => handleUserClick(user)}
              className="w-full text-left px-4 py-3 hover:bg-blue-50 border-b border-gray-100 last:border-b-0 transition-colors"
            >
              <div className="font-medium text-gray-900">{user.displayName}</div>
              <div className="text-sm text-gray-600">{user.mail || user.userPrincipalName}</div>
              {user.jobTitle && (
                <div className="text-xs text-gray-500 mt-1">{user.jobTitle}</div>
              )}
            </button>
          ))}
        </div>
      )}

      {/* No results message */}
      {showDropdown && !isSearching && results.length === 0 && debouncedSearchTerm.length >= 2 && (
        <div className="absolute z-10 w-full mt-1 bg-white border border-gray-300 rounded-lg shadow-lg p-4 text-sm text-gray-500">
          No users found matching "{debouncedSearchTerm}"
        </div>
      )}
    </div>
  );
};
