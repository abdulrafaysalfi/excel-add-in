import axios from "axios";

const instance = axios.create({
    baseURL: "https://ars-apis.herokuapp.com/api/",
});

export default instance;
