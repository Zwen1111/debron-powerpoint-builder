# Stage 1: Build
FROM --platform=$BUILDPLATFORM mcr.microsoft.com/dotnet/sdk:9.0 AS build
ARG BUILD_CONFIGURATION=Release
ARG TARGETPLATFORM
WORKDIR /src

# Kopieer alleen csproj en restore dependencies
COPY ["DeBron.PowerPoint.Builder/DeBron.PowerPoint.Builder.csproj", "DeBron.PowerPoint.Builder/"]
RUN dotnet restore "DeBron.PowerPoint.Builder/DeBron.PowerPoint.Builder.csproj"

# Kopieer de rest van de code
COPY . .
WORKDIR "/src/DeBron.PowerPoint.Builder"
RUN dotnet publish "DeBron.PowerPoint.Builder.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

# Stage 2: Runtime
FROM mcr.microsoft.com/dotnet/aspnet:9.0 AS final
WORKDIR /app
COPY --from=build /app/publish .
EXPOSE 8080
ENTRYPOINT ["dotnet", "DeBron.PowerPoint.Builder.dll"]
