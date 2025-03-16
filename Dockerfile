FROM node:20-bullseye AS base

# The "dependencies" stage
# It's good to install dependencies in a seprate stage to be explicit about
# the files that make it into production stage to avoid image bloat
FROM base AS dependencies

# The Application Directory
WORKDIR /app

# Copy fields for package managment
COPY package.json package-lock.json ./

# Install dependencies
RUN npm install


# The final image
FROM base AS production

# Create a group and a non-root user to run the app
RUN groupadd --gid 1001 "kleo"
RUN useradd --uid 1001 --create-home --shell /bin/bash --groups "kleo" "mottokuskjar"

# The Application Directory
WORKDIR /app

# Copy the dependencies from the "dependencies" stage
COPY --from=dependencies --chown=mottokuskjar:kleo /app/node_modules ./node_modules

# Copy the rest of the application
COPY --chown=mottokuskjar:kleo . .

# Enable production mode
ENV NODE_ENV=production

# Confiugre application port
ENV PORT=3000

# Expose the port
EXPOSE 3000

# Change the user
USER mottokusjakr:kleo