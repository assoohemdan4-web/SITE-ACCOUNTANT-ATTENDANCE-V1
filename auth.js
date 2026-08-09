const USERS = {
  accountant: { username: "accountant", password: "1234", role: "accountant", name: "Accountant" },
  employees: {
    "7636": { password: "1234", role: "employee", code: "7636", name: "Islam Mostafa Gamal El Din Abdallah", job: "Manager", site: "19 Dyarna" }
  }
};

const roleInput = document.getElementById("role");
document.querySelectorAll(".role-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".role-btn").forEach(x => x.classList.remove("active"));
    btn.classList.add("active");
    roleInput.value = btn.dataset.role;
    document.getElementById("identity").placeholder = btn.dataset.role === "accountant" ? "e.g. accountant" : "e.g. 7636";
  });
});
document.getElementById("togglePassword").addEventListener("click", () => {
  const input = document.getElementById("password");
  input.type = input.type === "password" ? "text" : "password";
});
document.getElementById("loginForm").addEventListener("submit", e => {
  e.preventDefault();
  const role = roleInput.value, id = document.getElementById("identity").value.trim(), pw = document.getElementById("password").value;
  let user = null;
  if (role === "accountant" && id === USERS.accountant.username && pw === USERS.accountant.password) user = USERS.accountant;
  if (role === "employee" && USERS.employees[id] && USERS.employees[id].password === pw) user = USERS.employees[id];
  if (!user) {
    document.getElementById("loginMessage").textContent = "Invalid credentials. Use the demo credentials shown below.";
    return;
  }
  localStorage.setItem("sa_session", JSON.stringify(user));
  location.href = user.role === "accountant" ? "accountant.html" : "employee.html";
});