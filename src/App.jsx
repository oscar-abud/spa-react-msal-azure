import { useState } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  // Estados para manejar los resultados de las peticiones HTTP
  const [resultado, setResult] = useState(null);
  const [statusHttp, setStatusHttp] = useState(null);
  const [cargando, setLoading] = useState(false);

  // Iniciar y cerrar sesión con MSAL
  const handleLogin = () => {
    instance.loginRedirect(loginRequest);
  };

  const handleLogout = () => {
    instance.logoutRedirect();
  };

  // Función genérica para consumir AWS API Gateway (con o sin token JWT)
  const llamarApiGateway = async (endpointUrl, enviarToken = true) => {
    setLoading(true);
    setResult(null);
    setStatusHttp(null);

    try {
      const headers = {};

      if (enviarToken) {
        // Solicitamos el token de acceso para la API backend
        const requestScope = {
          scopes: ["api://49ad4a51-39f7-48f3-8600-0f1368d19cf6/OT.Create"],
          account: accounts[0],
        };

        const tokenResponse = await instance
          .acquireTokenSilent(requestScope)
          .catch(async () => {
            return await instance.acquireTokenPopup(requestScope);
          });

        // Adjuntamos el token Bearer en la cabecera HTTP
        headers["Authorization"] = `Bearer ${tokenResponse.accessToken}`;
      }

      // Hacemos la petición al API Gateway
      const res = await fetch(endpointUrl, {
        method: "GET",
        headers,
      });

      setStatusHttp(res.status);
      const json = await res.json();
      setResult(json);
    } catch (err) {
      setResult({ error: err.message });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div
      style={{
        maxWidth: "800px",
        margin: "0 auto",
        padding: "20px",
        fontFamily: "Arial, sans-serif",
      }}
    >
      <h1>☁️ Solución Cloud Native: React + MSAL + AWS API Gateway</h1>

      {/* SECCIÓN DE AUTENTICACIÓN */}
      <header
        style={{
          padding: "15px",
          backgroundColor: "#f4f4f9",
          borderRadius: "8px",
          marginBottom: "20px",
        }}
      >
        {isAuthenticated ? (
          <div>
            <p>
              🟢 <strong>Estado:</strong> Sesión Activa
            </p>
            <p>
              <strong>Usuario:</strong> {accounts[0]?.name} (
              {accounts[0]?.username})
            </p>
            <button
              onClick={handleLogout}
              style={{
                backgroundColor: "#dc3545",
                color: "white",
                border: "none",
                padding: "10px 15px",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              Cerrar Sesión
            </button>
          </div>
        ) : (
          <div>
            <p>
              🔴 <strong>Estado:</strong> No autenticado
            </p>
            <button
              onClick={handleLogin}
              style={{
                backgroundColor: "#007bff",
                color: "white",
                border: "none",
                padding: "10px 15px",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              Iniciar Sesión con Microsoft
            </button>
          </div>
        )}
      </header>

      {/* SECCIÓN DE PRUEBAS DE ENDPOINTS DE AWS */}
      {isAuthenticated ? (
        <main
          style={{
            border: "1px solid #ddd",
            padding: "20px",
            borderRadius: "8px",
          }}
        >
          <h2>🧪 Pruebas de API Gateway con Autorizador JWT</h2>
          <p>
            Usa los botones para comprobar cómo el API Gateway bloquea (401) o
            permite (200) el acceso según la presencia del Token JWT:
          </p>

          {/* PRUEBAS RUTA V1 */}
          <div style={{ marginBottom: "15px" }}>
            <h3>
              📌 Versión 1: <code>/v1/datos</code>
            </h3>
            <button
              onClick={() =>
                llamarApiGateway(
                  "https://y7irt8jh22.execute-api.us-east-1.amazonaws.com/v1/datos",
                  false,
                )
              }
              disabled={cargando}
              style={{ marginRight: "10px", padding: "8px 12px" }}
            >
              Probar v1 (SIN Token)
            </button>
            <button
              onClick={() =>
                llamarApiGateway(
                  "https://y7irt8jh22.execute-api.us-east-1.amazonaws.com/v1/datos",
                  true,
                )
              }
              disabled={cargando}
              style={{
                backgroundColor: "#28a745",
                color: "white",
                border: "none",
                padding: "8px 12px",
                borderRadius: "4px",
                cursor: "pointer",
              }}
            >
              Probar v1 (CON Token JWT)
            </button>
          </div>

          {/* PRUEBAS RUTA V2 */}
          <div style={{ marginBottom: "20px" }}>
            <h3>
              📌 Versión 2: <code>/v2/datos</code>
            </h3>
            <button
              onClick={() =>
                llamarApiGateway(
                  "https://y7irt8jh22.execute-api.us-east-1.amazonaws.com/v2/datos",
                  false,
                )
              }
              disabled={cargando}
              style={{ marginRight: "10px", padding: "8px 12px" }}
            >
              Probar v2 (SIN Token)
            </button>
            <button
              onClick={() =>
                llamarApiGateway(
                  "https://y7irt8jh22.execute-api.us-east-1.amazonaws.com/v2/datos",
                  true,
                )
              }
              disabled={cargando}
              style={{
                backgroundColor: "#007bff",
                color: "white",
                border: "none",
                padding: "8px 12px",
                borderRadius: "4px",
                cursor: "pointer",
              }}
            >
              Probar v2 (CON Token JWT)
            </button>
          </div>

          {/* RESULTADOS HTTP */}
          {cargando && <p>⌛ Consultando endpoint...</p>}

          {statusHttp && (
            <div style={{ margin: "15px 0" }}>
              <strong>Código de Estado HTTP: </strong>
              <span
                style={{
                  color: statusHttp === 200 ? "green" : "red",
                  fontWeight: "bold",
                  fontSize: "1.1em",
                }}
              >
                {statusHttp}{" "}
                {statusHttp === 401
                  ? "🚫 Unauthorized (Acceso denegado por API Gateway)"
                  : statusHttp === 200
                    ? "✅ 200 OK (Token verificado exitosamente)"
                    : ""}
              </span>
            </div>
          )}

          {resultado && (
            <pre
              style={{
                backgroundColor: "#272822",
                color: "#f8f8f2",
                padding: "15px",
                borderRadius: "5px",
                overflowX: "auto",
              }}
            >
              {JSON.stringify(resultado, null, 2)}
            </pre>
          )}
        </main>
      ) : (
        <p style={{ color: "#666" }}>
          Inicia sesión arriba para desbloquear las pruebas de consumo hacia los
          microservicios.
        </p>
      )}
    </div>
  );
}

export default App;
