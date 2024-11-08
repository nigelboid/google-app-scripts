/**
 * GetRefreshTokenSchwab()
 *
 * Obtain a valid refresh token for Schwab API queries
 *
 */
function GetRefreshTokenSchwab(sheetID, verbose)
{
  const refreshTokenTTLOffsetDaysDefault = 7;
  const refreshTokenStaleLoggingThrottle = 60 * 60 * 24;
  const currentTime = new Date();
  const refreshTokenCopy = GetValueByName(sheetID, "ParameterSchwabTokenRefreshSaved", verbose);
  var refreshTokenExpirationTime = null;
  var refreshToken = GetValueByName(sheetID, "ParameterSchwabTokenRefresh", verbose);

  if (!refreshToken)
  {
    // Missing refresh token
    LogThrottled(sheetID, `Missing refresh token <${refreshToken}> -- obtain a new one!!!`, verbose, refreshTokenStaleLoggingThrottle);
  }
  else if (refreshToken != refreshTokenCopy)
  {
    // Looks like we have a new refresh token -- save a copy and update expiration time
    var refreshTokenTTLOffsetDays = GetValueByName(sheetID, "ParameterSchwabTokenRefreshTTL", verbose);
    SetValueByName(sheetID, "ParameterSchwabTokenRefreshSaved", refreshToken, verbose);

    if (!refreshTokenTTLOffsetDays)
    {
      // No value obtained for refresh token TTL -- use default
      refreshTokenTTLOffsetDays = refreshTokenTTLOffsetDaysDefault;
    }

    refreshTokenExpirationTime = currentTime;
    refreshTokenExpirationTime.setDate(refreshTokenExpirationTime.getDate() + refreshTokenTTLOffsetDays);
    SetValueByName(sheetID, "ParameterSchwabTokenRefreshTimeStamp", refreshTokenExpirationTime, verbose);
  }
  else
  {
    // The refresh token has not changed -- get its expiration time stamp
    refreshTokenExpirationTime = GetValueByName(sheetID, "ParameterSchwabTokenRefreshTimeStamp", verbose);
  }
  
  LogVerbose(`Refresh token: ${refreshToken}`, verbose);
  LogVerbose(`Refresh token expiration: ${refreshTokenExpirationTime}`, verbose);
    
  if (currentTime > refreshTokenExpirationTime)
  {
    // Current refresh token has also expired or has no expiration value -- report and invalidate
    LogThrottled(sheetID, "Refresh token has gone stale -- obtain a new one!!!", verbose, refreshTokenStaleLoggingThrottle);

    refreshToken = null;
  }

  return refreshToken;
};