FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY NPI_ACS_Web/NPI_ACS_Web.csproj NPI_ACS_Web/
RUN dotnet restore NPI_ACS_Web/NPI_ACS_Web.csproj

COPY . .
WORKDIR /src/NPI_ACS_Web
RUN dotnet publish -c Release -o /app

FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /app .

ENTRYPOINT ["dotnet", "NPI_ACS_Web.dll"]