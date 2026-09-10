import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const handleLogin = () => {
    instance.loginRedirect(loginRequest);
  };

  const handleLogout = () => {
    instance.logoutRedirect();
  };

  return (
    <div style={{ padding: "20px", fontFamily: "sans-serif" }}>
      <h1>Autenticación con MSAL en React + Vite</h1>
      {isAuthenticated ? (
        <div>
          <p>
            Bienvenido, <strong>{accounts[0]?.name}</strong> (
            {accounts[0]?.username})
          </p>
          <button onClick={handleLogout}>Cerrar sesión</button>
        </div>
      ) : (
        <div>
          <p>No has iniciado sesión.</p>
          <button onClick={handleLogin}>Iniciar sesión con Microsoft</button>
        </div>
      )}
    </div>
  );
}

export default App;
