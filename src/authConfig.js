export const msalConfig = {
  auth: {
    clientId: "b79b987c-d9bd-4b63-a995-40b4cff0113a", // Pega tu Client ID aquí
    authority: "https://login.microsoftonline.com/635fce4d-bfca-4393-994d-91819d9056e1", // Pega tu Tenant ID aquí
    redirectUri: "http://localhost:5173",
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: ["User.Read"],
};
